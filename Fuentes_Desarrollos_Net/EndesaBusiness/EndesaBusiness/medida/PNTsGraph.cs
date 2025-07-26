using Azure.Identity;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class PNTsGraph
    {
        private const string clientId = "53965a60-a073-4e21-899a-d8f13305c171";        
        private const string clientSecret = "53965a60-a073-4e21-899a-d8f13305c171";
        private const string tenantId = "d539d4bf-5610-471a-afc2-1c76685cfefa";
        private const string userId = "rsiope.gma@enel.com";
        GraphServiceClient graphClient;
        logs.Log ficheroLog;

        public PNTsGraph() 
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "PNTsGraph");
            graphClient = Connectar_client_credential();
        }

        private GraphServiceClient Connectar_client_credential()
        {
            try
            {
                // The client credentials flow requires that you request the
                // /.default scope, and preconfigure your permissions on the
                // app registration in Azure. An administrator must grant consent
                // to those permissions beforehand.
                var scopes = new[] { "https://graph.microsoft.com/.default" };
                //var scopes = new[] { "https://graph.microsoft.com/mail.send" };
                //var scopes = new[] { "email" };

                // Multi-tenant apps can use "common",
                // single-tenant apps must use the tenant ID from the Azure portal
                var tenantId = "d539d4bf-5610-471a-afc2-1c76685cfefa";

                // Values from app registration
                var clientId = "53965a60-a073-4e21-899a-d8f13305c171";
                var clientSecret = "dIW8Q~g3vdKpGEHow95dGGrstv8AxeAEAE1vdby2";

                // using Azure.Identity;
                var options = new TokenCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
                };

                // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
                var clientSecretCredential = new ClientSecretCredential(
                    tenantId, clientId, clientSecret, options);

                var graphClient = new GraphServiceClient(clientSecretCredential, scopes);
                //txtInfo.AppendText(graphClient.Me.Authentication.Client.AuthenticationProvider + Environment.NewLine);
                Console.WriteLine(graphClient.Me.Authentication.Client.AuthenticationProvider);

                return graphClient;
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Connectar_client_credential: " + e.Message);
                //txtInfo.AppendText(e.Message + Environment.NewLine);
                Console.WriteLine(e.Message);
                //txtInfo.AppendText(e.InnerException.Message + Environment.NewLine);
                Console.WriteLine("Connectar_client_credential" + e.InnerException.Message);
                ficheroLog.AddError(e.InnerException.Message);
                throw;
            }
        }

        public void RecorreBuzon()
        {
            //var messages = await graphClient.Users[userId].MailFolders.Inbox.Messages.Request().GetAsync();
            graphClient.Me.MailFolders.Request().GetAsync().Wait(); 
                       


           
        }


    }
}
