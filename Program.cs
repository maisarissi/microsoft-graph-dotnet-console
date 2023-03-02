using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

public class Program
{
    public static async Task Main(string[] args)
    {
        var settings = Settings.LoadSettings();
        string[] scopes = {"User.Read", "Mail.Read"};
        GraphServiceClient graphClient = new Authentication().getGraphClient(scopes, settings);

        User? me = await graphClient.Me.GetAsync(
            requestConfig => {
                //You can also request specific properties
                requestConfig.QueryParameters.Select = new string[] {"id", "displayName", "mail", "userPrincipalName", "createdDateTime"};
            }
        );

        Console.WriteLine($"Hello {me?.GivenName}!"); //givenName is null because this properties wasn't requested
        Console.WriteLine($"$Hello {me.DisplayName}! Your email is {me?.Mail}");

        //Get the first page of messages - the default page size is 10
        //More info: https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=csharp#request
        var firstPage = await graphClient.Me.Messages.GetAsync(
            /*You can also change the page size
            requestConfig => {
                requestConfig.QueryParameters.Top = 5;
            }*/
        );

        Console.WriteLine($"Fetched {firstPage.Value.Count} messages via default request");

        var messagesCollected = new List<Message>();
        
        //We can leverage the PageIterator to iterate over the pages
        var pageIterator = PageIterator<Message, MessageCollectionResponse>.CreatePageIterator(
            graphClient, 
            firstPage,
            message =>
            {
                messagesCollected.Add(message);
                return true;
            },//Per item callback
            request =>
            {
                Console.WriteLine($"Requesting new page with url {request.URI.OriginalString}");
                return request;
            }//Per request/page callback to reconfigure the request
        );

        //Then iterate over the pages
        await pageIterator.IterateAsync();

        // Get the messages data;
        Console.WriteLine($"Fetched {messagesCollected.Count} messages via page iterator");
        Console.WriteLine("-----------Done with paged requests-----------");
    }
}