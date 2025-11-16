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

// // Test the WhoAmI tool directly
// var serviceProvider = builder.Build().Services;
// var orgService = serviceProvider.GetRequiredService<IOrganizationService>();
// var whoAmIResult = DataverseTool.WhoAmI(orgService);
// Console.WriteLine(whoAmIResult);

builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly();
await builder.Build().RunAsync();

[McpServerToolType]
public static class DataverseTool
{
    [McpServerTool, Description("Executes an WhoAmI request aginst Dataverse and returns the result as a JSON string.")]
    public static string WhoAmI(IOrganizationService orgService)
    {
        try
        {
            WhoAmIRequest req = new WhoAmIRequest();

            var whoAmIResult = orgService.Execute(req);

            return Newtonsoft.Json.JsonConvert.SerializeObject(whoAmIResult);
        }
        catch (Exception err)
        {
            Console.Error.WriteLine(err.ToString());

            return err.ToString();
        }
    }

    [McpServerTool, Description("Executes an FetchXML request using the supplied expression that needs to be a valid FetchXml expression. Returns the result as a JSON string. If the request fails, the response will be prepended with [ERROR] and the error should be presented to the user.")]
    public static string ExecuteFetch(string fetchXmlRequest, IOrganizationService orgService)
    {
        try
        {
            FetchExpression fetchExpression = new FetchExpression(fetchXmlRequest);
            EntityCollection result = orgService.RetrieveMultiple(fetchExpression);

            return Newtonsoft.Json.JsonConvert.SerializeObject(result);
        }
        catch (Exception err)
        {
            var errorString = "[ERROR] " + err.ToString();
            Console.Error.WriteLine(err.ToString());

            return errorString;
        }
    }

    [McpServerTool, Description("Creates a Speaker.")]
    public static string CreateSpeaker(string firstname, string lastname, IOrganizationService orgService)
    {
        try
        {
            Entity contact = new Entity("contact");
            contact["firstname"] = firstname;
            contact["lastname"] = lastname;

            Guid contactId = orgService.Create(contact);

            return $"Contact created successfully with ID: {contactId}";
        }
        catch (Exception err)
        {
            var errorString = "[ERROR] " + err.ToString();
            Console.Error.WriteLine(err.ToString());

            return errorString;
        }
    }

    [McpServerTool, Description("Creates an Event.")]
    public static string CreateEvent(string eventName, string location, DateTime eventDate, IOrganizationService orgService)
    {
        try
        {
            Entity eventEntity = new Entity("new_event");
            eventEntity["cr5ec_eventname"] = eventName;
            eventEntity["new_location"] = location;
            eventEntity["cr5ec_eventdate"] = eventDate;

            Guid eventId = orgService.Create(eventEntity);

            return $"Event created successfully with ID: {eventId}";
        }
        catch (Exception err)
        {
            var errorString = "[ERROR] " + err.ToString();
            Console.Error.WriteLine(err.ToString());

            return errorString;
        }
    }

    [McpServerTool, Description("Adds a speaker to an event.")]
    public static string AddSpeakerToEvent(Guid eventId, Guid speakerId, IOrganizationService orgService)
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

            orgService.Execute(request);

            return $"Speaker {speakerId} successfully added to event {eventId}";
        }
        catch (Exception err)
        {
            var errorString = "[ERROR] " + err.ToString();
            Console.Error.WriteLine(err.ToString());

            return errorString;
        }
    }

    [McpServerTool, Description("Updates the biography of a speaker.")]
    public static string UpdateSpeakerBiography(Guid speakerId, string biography, IOrganizationService orgService)
    {
        try
        {
            Entity contact = new Entity("contact", speakerId);
            contact["cr5ec_biography"] = biography;

            UpdateRequest request = new UpdateRequest()
            {
                Target = contact
            };

            orgService.Execute(request);

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






