using Azure.Identity;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Graph;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.office365
{
    public class SharePoint_Teams
    {

        logs.Log ficheroLog;
        GraphServiceClient graphClient;
        public SharePoint_Teams()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "SharePoint_Teams");
        }


        public void Prueba()
        {
            //using (var client = new HttpClient())
            //{
            //    var url = "https://graph.microsoft.com/v1.0" + $"/drives/{driveID}/items/{folderId}:/{originalFileName}:/createUploadSession";
            //    client.DefaultRequestHeaders.Add("Authorization", "Bearer " 
            //        + graphClient.Me.Authentication.Client.AuthenticationProvider);

            //    var sessionResponse = client.PostAsync(apiUrl, null).Result.Content.ReadAsStringAsync().Result;

            //    byte[] sContents = System.IO.File.ReadAllBytes(filePath);
            //    var uploadSession = JsonConvert.DeserializeObject<UploadSessionResponse>(sessionResponse);
            //    string response = UploadFileBySession(uploadSession.uploadUrl, sContents);
            //}
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


        //private string UploadFileBySession(string url, byte[] file)
        //{
        //    //int fragSize = 1024 * 1024 * 4;
        //    //var arrayBatches = ByteArrayIntoBatches(file, fragSize);
        //    //int start = 0;
        //    //string response = "";

        //    //foreach (var byteArray in arrayBatches)
        //    //{
        //    //    int byteArrayLength = byteArray.Length;
        //    //    var contentRange = " bytes " + start + "-" + (start + (byteArrayLength - 1)) + "/" + file.Length;

        //    //    using (var client = new HttpClient())
        //    //    {
        //    //        var content = new ByteArrayContent(byteArray);
        //    //        content.Headers.Add("Content-Length", byteArrayLength.ToProperString());
        //    //        content.Headers.Add("Content-Range", contentRange);

        //    //        response = client.PutAsync(url, content).Result.Content.ReadAsStringAsync().Result;
        //    //    }

        //    //    start = start + byteArrayLength;
        //    //}
        //    //return response;
        //}

        internal IEnumerable<byte[]> ByteArrayIntoBatches(byte[] bArray, int intBufforLengt)
        {
            int bArrayLenght = bArray.Length;
            byte[] bReturn = null;

            int i = 0;
            for (; bArrayLenght > (i + 1) * intBufforLengt; i++)
            {
                bReturn = new byte[intBufforLengt];
                Array.Copy(bArray, i * intBufforLengt, bReturn, 0, intBufforLengt);
                yield return bReturn;
            }

            int intBufforLeft = bArrayLenght - i * intBufforLengt;
            if (intBufforLeft > 0)
            {
                bReturn = new byte[intBufforLeft];
                Array.Copy(bArray, i * intBufforLengt, bReturn, 0, intBufforLeft);
                yield return bReturn;
            }
        }
    }
}
