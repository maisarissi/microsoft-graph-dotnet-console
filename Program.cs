using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

public class Program
{
    public static async Task Main(string[] args)
    {
        var settings = Settings.LoadSettings();
        string[] scopes = {"User.Read", "Mail.Read", "Mail.ReadWrite", "Mail.Send"};
        GraphServiceClient graphClient = new Authentication().getGraphClient(scopes, settings);

        User? me = await graphClient.Me.GetAsync(
            requestConfig => {
                //You can also request specific properties
                requestConfig.QueryParameters.Select = new string[] {"id", "displayName", "mail", "userPrincipalName", "createdDateTime"};
            }
        );

        Console.WriteLine($"Hello {me?.GivenName}!"); //givenName is null because this properties wasn't requested
        Console.WriteLine($"$Hello {me?.DisplayName}! Your email is {me?.Mail}");

        //Get the first page of messages - the default page size is 10
        //More info: https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=csharp#request
        var firstPage = await graphClient.Me.Messages.GetAsync(
            /*You can also change the page size
            requestConfig => {
                requestConfig.QueryParameters.Top = 5;
            }*/
        );

        Console.WriteLine($"Fetched {firstPage?.Value?.Count} messages via default request");

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
        Console.WriteLine($"First message subject: {messagesCollected[0].Subject}");
        Console.WriteLine($"Last message subject: {messagesCollected[messagesCollected.Count - 1].Subject}");
        Console.WriteLine("-----------Done with paged requests-----------");
        
        //Sending a message
        var message = new Message
        {
            Subject = "Hello from Graph!",
            Body = new ItemBody{
                Content = "Hello from Graph!",
                ContentType = BodyType.Text
            },
            ToRecipients = new List<Recipient>{
                new Recipient{
                    EmailAddress = new EmailAddress{
                        Address = "admin@M365x18467905.onmicrosoft.com"
                    }
                }
            }
        };

        var saveToSentItems = true;

        var body = new SendMailPostRequestBody 
        {
            Message = message,
            SaveToSentItems = saveToSentItems
        };

        try
        {
            await graphClient.Me.SendMail.PostAsync(body);
            Console.WriteLine("Message sent!");
        }
        catch (ODataError odataError)
        {
            Console.WriteLine(odataError.Error.Code);
            Console.WriteLine(odataError.Error.Message);
            throw;
        }

        var lastMessage = await graphClient.Me.Messages.GetAsync(
            requestConfig => {
                requestConfig.QueryParameters.Top = 1;
            }
        );

        Console.WriteLine("lastMessage: " + lastMessage.Value[0].Subject);
        Console.WriteLine("lastMessage: " + lastMessage.Value[0].Id);

        WaitCallback waitCallback = new WaitCallback((state) => {
            Thread.Sleep(5000);
            Console.WriteLine("WaitCallback: " + state);
        });

        try
        {
            await graphClient.Me.Messages[lastMessage.Value[0].Id].DeleteAsync();
            Console.WriteLine("Message deleted!");
        }
        catch (ODataError odataError)
        {
            Console.WriteLine(odataError.Error.Code);
            Console.WriteLine(odataError.Error.Message);
            throw;
        }

        var lastMessageAfterDeleting = await graphClient.Me.Messages.GetAsync(
            requestConfig => {
                requestConfig.QueryParameters.Top = 1;
            }
        );

        Console.WriteLine("lastMessageAfterDeleting: " + lastMessageAfterDeleting.Value[0].Subject);

    }
}