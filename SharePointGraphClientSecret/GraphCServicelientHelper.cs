using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph.ExternalConnectors;
using Azure.Identity;
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
            var clientId = "f11cfc01-4382-4b70-9996-b9494a15d133";
            var clientSecret = "nzx8Q~iFItIcm5rfzcxu3SjpQZzNin7kUHMMhbcM";

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
