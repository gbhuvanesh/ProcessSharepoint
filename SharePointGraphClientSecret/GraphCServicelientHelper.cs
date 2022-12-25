using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.ExternalConnectors;

namespace SharePointGraphClientSecret
{
    internal class GraphCServicelientHelper
    {
        public GraphCServicelientHelper() { }

        public GraphServiceClient GetGraphClient()
        {

            var scopes = new string[] { "https://graph.microsoft.com/.default" };
            var tenantId = "6716eb25-dab8-4305-a4a4-ab061d87f701";

            // Configure the MSAL client as a confidential client
            var confidentialClient = ConfidentialClientApplicationBuilder
                            .Create("c6ec4c46-12c7-48db-b314-770081f65b4a")
             .WithAuthority($"https://login.microsoftonline.com/{tenantId}/v2.0")
                            .WithClientSecret("sfGhzJCBhlGR11oyKob7RXWaU8BDBJtVZEjTW9Ri7I4")
                            .Build();

            // Build the Microsoft Graph client. As the authentication provider, set an async lambda
            // which uses the MSAL client to obtain an app-only access token to Microsoft Graph,
            // and inserts this access token in the Authorization header of each API request. 

            return new GraphServiceClient(
                
                new DelegateAuthenticationProvider(async (requestMessage) =>
            {

                // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                var authResult = await confidentialClient
                         .AcquireTokenForClient(scopes)
                         .ExecuteAsync();

                // Add the access token in the Authorization header of the API request.
                requestMessage.Headers.Authorization =
                                   new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
            })
                );
        
        }
    }
}
