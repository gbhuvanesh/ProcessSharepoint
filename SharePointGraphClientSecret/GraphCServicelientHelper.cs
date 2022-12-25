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
                            .Create("13985f93-53cb-4db8-b814-8e02a2e80e80 ")
             .WithAuthority($"https://login.microsoftonline.com/{tenantId}/v2.0")
                            .WithClientSecret("qSU8Q~bA8Pqf2Zk1bjTY1zUVuwv4bxH0giRAodlX")
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
