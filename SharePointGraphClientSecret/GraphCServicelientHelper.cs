using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph.ExternalConnectors;
using Azure.Identity;.
namespace SharePointGraphClientSecret
{
    internal class GraphCServicelientHelper
    {
        public GraphCServicelientHelper() { }

        public GraphServiceClient GetGraphClient()
        {
            // The client credentials flow requires that you request the
            // /.default scope, and preconfigure your permissions on the
            // app registration in Azure. An administrator must grant consent
            // to those permissions beforehand.
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            // Multi-tenant apps can use "common",
            // single-tenant apps must use the tenant ID from the Azure portal
            var tenantId = "6716eb25-dab8-4305-a4a4-ab061d87f701";

            // Values from app registration
            var clientId = "c6ec4c46-12c7-48db-b314-770081f65b4a";
            var clientSecret = "sfGhzJCBhlGR11oyKob7RXWaU8BDBJtVZEjTW9Ri7I4";

            // using Azure.Identity;
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, options);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            return graphClient;
        }
    }
}
