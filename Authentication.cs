using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Authentication;
using Microsoft.Kiota.Http.HttpClientLibrary.Middleware;
using Microsoft.Kiota.Http.HttpClientLibrary.Middleware.Options;

public class Authentication {

    public GraphServiceClient getGraphClient(string[] scopes, Settings settings){
        
        InteractiveBrowserCredentialOptions interactiveBrowserCredentialOptions = new InteractiveBrowserCredentialOptions
        {
            ClientId = settings.ClientId
        };

        //Getting the token credential using my own clientId
        InteractiveBrowserCredential tokenCredential = new InteractiveBrowserCredential(interactiveBrowserCredentialOptions);

        //This creates a client with default handlers
        GraphServiceClient graphClient = new GraphServiceClient(tokenCredential, scopes);


        //You can also customize your client by adding your own middlewares
        //The steps below show how to add a custom middleware to the default set of middlewares
        var authProvider = new AzureIdentityAuthenticationProvider(tokenCredential, scopes);

        //The Microsoft Graph client library configures a default set of middlewares
        var handlers = GraphClientFactory.CreateDefaultHandlers();

        //Add a custom handler to the list of handlers
        handlers.Add(new ChaosHandler(new ChaosHandlerOption()
        {
            ChaosPercentLevel = 50
        }));

        var httpClient = GraphClientFactory.Create(handlers);

        var customGraphClient = new GraphServiceClient(httpClient, authProvider);

        //Here I'm returning the default client, but you can return the custom one
        return graphClient;
    }
}