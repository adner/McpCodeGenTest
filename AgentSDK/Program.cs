using Microsoft.Agents.AI;
using OpenAI;

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.Extensions.AI;
using System.ComponentModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.PowerPlatform.Dataverse.Client;
using MicrosoftAgentFramework.Utilities.Extensions;
using ScriptRunnerLib;
using System.Text;

var configuration = new ConfigurationBuilder()
    .AddJsonFile("appsettings.Development.json")
    .Build();

var apiKey = configuration["OpenAI:ApiKey"] ?? throw new InvalidOperationException("OpenAI:ApiKey is missing from configuration");

var chatClient = new OpenAIClient(apiKey).GetChatClient("gpt-5.1");

string url = configuration["Dataverse:Url"] ?? throw new InvalidOperationException("Dataverse:Url is missing from configuration");
string clientId = configuration["Dataverse:ClientId"] ?? throw new InvalidOperationException("Dataverse:ClientId is missing from configuration");
string secret = configuration["Dataverse:Secret"] ?? throw new InvalidOperationException("Dataverse:Secret is missing from configuration");
string connectionString = $@"
    AuthType = ClientSecret;
    Url = {url};
    ClientId = {clientId};
    Secret = {secret}";

var serviceClient = new ServiceClient(connectionString);

DataverseTools dataverseTools = new DataverseTools(serviceClient);

var scriptRunner = new ScriptRunner<DataverseTools>(dataverseTools);

// Read agent instructions from file
string pastEvents = File.ReadAllText("past_events.txt");
string codeGenInstructions = File.ReadAllText("agent_instructions_codegen.md");

AIAgent toolCallingAgent = chatClient.CreateAIAgent( instructions: "You are a helpful agent that can create Speakers and Events in Dataverse by using the tools CreateSpeaker and CreateEvent.",
        tools:
        [
            AIFunctionFactory.Create(dataverseTools.CreateSpeaker),
            AIFunctionFactory.Create(dataverseTools.CreateEvent)
        ]).AsBuilder()
    .Use(FunctionCallMiddleware) //Middleware
    .Build();

AIAgent codeGentAgent = chatClient.CreateAIAgent(instructions: codeGenInstructions,
        tools:
        [
           AIFunctionFactory.Create(scriptRunner.RunScript)
        ]).AsBuilder()
    .Use(FunctionCallMiddleware) //Middleware
    .Build();
;

var prompt = "Please go through these past events and create the Events and Speakers in Dataverse. After creating each Event or Speaker, make sure that it returns 'OK', before continuing to the next one. When done with all, just say 'Done!'" + pastEvents;
//var prompt = "Please go through these past events and create the Events and Speakers in Dataverse. When done with all, just say 'Done!'" + pastEvents;

System.Console.WriteLine("Running tool calling agent...");
var toolCallingSw = System.Diagnostics.Stopwatch.StartNew();
await RunAgentAndDisplayUsage(toolCallingAgent, prompt);
toolCallingSw.Stop();
Console.WriteLine($"\nTool Calling Agent Elapsed Time: {toolCallingSw.ElapsedMilliseconds}ms ({toolCallingSw.Elapsed.TotalSeconds:F2}s)");

System.Console.WriteLine("\nRunning code gen agent...");
var codeGenSw = System.Diagnostics.Stopwatch.StartNew();
await RunAgentAndDisplayUsage(codeGentAgent, prompt);
codeGenSw.Stop();
Console.WriteLine($"\nCode Gen Agent Elapsed Time: {codeGenSw.ElapsedMilliseconds}ms ({codeGenSw.Elapsed.TotalSeconds:F2}s)");

static async Task<AgentRunResponse> RunAgentAndDisplayUsage(AIAgent agent, string prompt)
{
    List<AgentRunResponseUpdate> updates = [];

    await foreach (AgentRunResponseUpdate update in agent.RunStreamingAsync(prompt))
    {
        updates.Add(update);
        Console.Write(update);
    }

    AgentRunResponse collectedResponseFromStreaming = updates.ToAgentRunResponse();

    Console.WriteLine($"\n\n- Input Tokens (Streaming): {collectedResponseFromStreaming.Usage?.InputTokenCount}");
    Console.WriteLine($"- Output Tokens (Streaming): {collectedResponseFromStreaming.Usage?.OutputTokenCount} " +
                            $"({collectedResponseFromStreaming.Usage?.GetOutputTokensUsedForReasoning()} was used for reasoning)");
    
    Console.Write($"- Total tokens (Streaming): {collectedResponseFromStreaming.Usage?.InputTokenCount + collectedResponseFromStreaming.Usage?.OutputTokenCount + collectedResponseFromStreaming.Usage?.GetOutputTokensUsedForReasoning()}" );

    return collectedResponseFromStreaming;
}

async ValueTask<object?> FunctionCallMiddleware(AIAgent callingAgent, FunctionInvocationContext context, Func<FunctionInvocationContext, CancellationToken, ValueTask<object?>> next, CancellationToken cancellationToken)
{
    StringBuilder functionCallDetails = new();
    functionCallDetails.Append($"- Tool Call: '{context.Function.Name}'");
    if (context.Arguments.Count > 0)
    {
        functionCallDetails.Append($" (Args: {string.Join(",", context.Arguments.Select(x => $"[{x.Key} = {x.Value}]"))}");
    }

    Console.WriteLine(functionCallDetails.ToString());

    return await next(context, cancellationToken);
}

public class DataverseTools
{   
    private IOrganizationService _orgService;
    public DataverseTools(IOrganizationService service)
    {
        _orgService = service;      
    }

    [Description("Executes an WhoAmI request aginst Dataverse and returns the result as a JSON string.")]
    public string WhoAmI()
    {
        try
        {
            WhoAmIRequest req = new WhoAmIRequest();

            var whoAmIResult = _orgService.Execute(req);

            return Newtonsoft.Json.JsonConvert.SerializeObject(whoAmIResult);
        }
        catch (Exception err)
        {
            Console.Error.WriteLine(err.ToString());

            return err.ToString();
        }
    }

    [Description("Executes an FetchXML request using the supplied expression that needs to be a valid FetchXml expression. Returns the result as a JSON string. If the request fails, the response will be prepended with [ERROR] and the error should be presented to the user.")]
    public string ExecuteFetch(string fetchXmlRequest)
    {
        try
        {
            FetchExpression fetchExpression = new FetchExpression(fetchXmlRequest);
            EntityCollection result = _orgService.RetrieveMultiple(fetchExpression);

            return Newtonsoft.Json.JsonConvert.SerializeObject(result);
        }
        catch (Exception err)
        {
            var errorString = "[ERROR] " + err.ToString();
            Console.Error.WriteLine(err.ToString());

            return errorString;
        }
    }

    [Description("Creates a Speaker.")]
    public string CreateSpeaker(string firstname, string lastname)
    {
        try
        {
            Entity contact = new Entity("contact");
            contact["firstname"] = firstname;
            contact["lastname"] = lastname;

            Guid contactId = _orgService.Create(contact);

            return $"OK";
        }
        catch (Exception err)
        {
            var errorString = "[ERROR] " + err.ToString();
            Console.Error.WriteLine(err.ToString());

            return errorString;
        }
    }

    [Description("Creates an Event.")]
    public string CreateEvent(string eventName, string location, DateTime eventDate)
    {
        try
        {
            Entity eventEntity = new Entity("new_event");
            eventEntity["cr5ec_eventname"] = eventName;
            eventEntity["new_location"] = location;
            eventEntity["cr5ec_eventdate"] = eventDate;

            Guid eventId = _orgService.Create(eventEntity);

            return $"OK";
        }
        catch (Exception err)
        {
            var errorString = "[ERROR] " + err.ToString();
            Console.Error.WriteLine(err.ToString());

            return errorString;
        }
    }

    [Description("Adds a speaker to an event.")]
    public string AddSpeakerToEvent(Guid eventId, Guid speakerId)
    {
        try
        {
            Relationship relationship = new Relationship("cr5ec_new_EventSpeakers");

            AssociateRequest request = new AssociateRequest()
            {
                Target = new EntityReference("new_event", eventId),
                RelatedEntities = new EntityReferenceCollection()
                {
                    new EntityReference("contact", speakerId)
                },
                Relationship = relationship
            };

           _orgService.Execute(request);

            return $"Speaker {speakerId} successfully added to event {eventId}";
        }
        catch (Exception err)
        {
            var errorString = "[ERROR] " + err.ToString();
            Console.Error.WriteLine(err.ToString());

            return errorString;
        }
    }

    [Description("Updates the biography of a speaker.")]
    public string UpdateSpeakerBiography(Guid speakerId, string biography)
    {
        try
        {
            Entity contact = new Entity("contact", speakerId);
            contact["cr5ec_biography"] = biography;

            UpdateRequest request = new UpdateRequest()
            {
                Target = contact
            };

            _orgService.Execute(request);

            return $"Speaker biography successfully updated for {speakerId}.";
        }
        catch (Exception err)
        {
            var errorString = "[ERROR] " + err.ToString();
            Console.Error.WriteLine(err.ToString());

            return errorString;
        }
    }
}

