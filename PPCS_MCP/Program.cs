using Microsoft.CodeAnalysis.Scripting;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using ModelContextProtocol.Server;
using System.ComponentModel;
using ScriptRunnerLib;

var builder = Host.CreateApplicationBuilder(args);

builder.Logging.AddConsole(consoleLogOptions =>
{
    // Configure all logs to go to stderr
    consoleLogOptions.LogToStandardErrorThreshold = LogLevel.Trace;
});

builder.Services.AddSingleton<IOrganizationService>(provider =>
{
    // TODO Enter your Dataverse environment's URL and logon info.
    string url = "[Url]";
    string connectionString = $@"
    AuthType = ClientSecret;
    Url = {url};
    ClientId = [ClientId];
    Secret = [ClientSecret]";

    return new ServiceClient(connectionString);
});

// Test the WhoAmI tool directly
// var serviceProvider = builder.Build().Services;
// var orgService = serviceProvider.GetRequiredService<IOrganizationService>();
// var codeGenTool = new CodeGenTool(orgService);

// var result = codeGenTool.RunScript("var user = WhoAmI();user");

builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly();
await builder.Build().RunAsync();

[McpServerToolType]
public class CodeGenTool
{
    private ScriptRunnerLib.ScriptRunner<DataverseTools> _scriptRunner;

    public CodeGenTool(IOrganizationService orgService)
    {
        _scriptRunner = new ScriptRunnerLib.ScriptRunner<DataverseTools>(new DataverseTools(orgService));
    }

    [McpServerTool, Description("Call this tool to execute C# code.")]
    public string RunScript(string code)
    {
        var result = _scriptRunner.RunScript(code).Result;

        return result;
    }
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

            return $"Contact created successfully with ID: {contactId}";
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

            return $"Event created successfully with ID: {eventId}";
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







