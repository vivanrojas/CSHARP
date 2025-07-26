using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using System.IO;


namespace EndesaBusiness.office365
{
    public class OAuth_Mail
    {
        string clientId;
        string appId;
        string clientSecret;
        string tenantId;
        GraphServiceClient graphServiceClient;
        logs.Log ficheroLog;
        public OAuth_Mail()
        {

            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "OAuth_Mail");
            graphServiceClient = Connectar_client_credential();
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

        public void DownloadFile()
        {
            graphServiceClient
                   .Sites["root"]
                   .SiteWithPath("/teams/myawesometeam")
                   .Request()
                   .GetAsync().Wait();


        }

        //public void UploadFile()
        //{
        //    var fileStream = System.IO.File.OpenRead(@"C:\Temp\contratacion\PS_AT_221007.zip");

        //    var uploadProps = new DriveItemUploadableProperties
        //    {
        //        AdditionalData = new Dictionary<string, object>
        //        {
        //            { "@microsoft.graph.conflictBehavior", "replace" }
        //         }
        //    };

        //    string itemPath = @"C:\Temp\contratacion\PS_AT_221007.zip";
        //    // Create the upload session
        //    // itemPath does not need to be a path to an existing item
        //    graphServiceClient.Me.Drive.Root
        //        .ItemWithPath(itemPath)
        //        .CreateUploadSession(uploadProps)
        //        .Request()
        //        .PostAsync().Wait();

        //    // Max slice size must be a multiple of 320 KiB
        //    int maxSliceSize = 320 * 1024;
        //    var fileUploadTask =
        //        new LargeFileUploadTask<DriveItem>(UploadSession, fileStream, maxSliceSize);

        //    var totalLength = fileStream.Length;
        //    // Create a callback that is invoked after each slice is uploaded
        //    IProgress<long> progress = new Progress<long>(prog => {
        //        Console.WriteLine($"Uploaded {prog} bytes of {totalLength} bytes");
        //    });

        //    try
        //    {
        //        // Upload the file
        //        var uploadResult = await fileUploadTask.UploadAsync(progress).Wait();

        //        Console.WriteLine(uploadResult.UploadSucceeded ?
        //            $"Upload complete, item ID: {uploadResult.ItemResponse.Id}" :
        //            "Upload failed");
        //    }
        //    catch (ServiceException ex)
        //    {
        //        Console.WriteLine($"Error uploading: {ex.ToString()}");
        //    }

        //}

        public string SendMail(string from, string to, string cc, string subject, string body)
        {
            try
            {
                               

                string[] toMail = to.Split(';');
                List<Recipient> toRecipients = new List<Recipient>();
                int i = 0;
                for (i = 0; i < toMail.Count(); i++)
                {
                    Recipient toRecipient = new Recipient();
                    EmailAddress toEmailAddress = new EmailAddress();

                    toEmailAddress.Address = toMail[i];
                    toRecipient.EmailAddress = toEmailAddress;
                    toRecipients.Add(toRecipient);
                }

                List<Recipient> ccRecipients = new List<Recipient>();
                if (!string.IsNullOrEmpty(cc))
                {
                    string[] ccMail = cc.Split(';');
                    int j = 0;
                    for (j = 0; j < ccMail.Count(); j++)
                    {
                        Recipient ccRecipient = new Recipient();
                        EmailAddress ccEmailAddress = new EmailAddress();

                        ccEmailAddress.Address = ccMail[j];
                        ccRecipient.EmailAddress = ccEmailAddress;
                        ccRecipients.Add(ccRecipient);
                    }
                }

                              



                var mailMessage = new Microsoft.Graph.Message
                {
                    Subject = subject,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = body
                    },
                    ToRecipients = toRecipients,
                    CcRecipients = ccRecipients
                    

                };

                ficheroLog.Add("from: " + from + " to: " + to + " subject: " + subject);

                
                graphServiceClient
                   .Users[from]
                    .SendMail(mailMessage, true)
                    .Request()
                    .PostAsync().Wait();


                ficheroLog.Add("Mail enviado correctamente");
                return "Email successfully sent.";


            }
            catch (Exception ex)
            {
                string cadena_error = ex.Message;
                ficheroLog.AddError("SendMail: " + ex.Message);
                if (ex.InnerException != null)
                {
                    Console.WriteLine("SendMail: " + ex.InnerException.Message);
                    cadena_error += Environment.NewLine + ex.InnerException.Message;
                    if (ex.InnerException.InnerException != null)
                    {
                        
                        cadena_error += Environment.NewLine + ex.InnerException.InnerException.Message;
                        ficheroLog.AddError("SendMail: " + ex.InnerException.InnerException.Message);
                    }
                }

                return cadena_error;
            }
        }


        public string SendMail(string from, string to, string cc, string subject, string body, string attachements)
        {
            try
            {


                string[] toMail = to.Split(';');
                List<Recipient> toRecipients = new List<Recipient>();                
                for (int i = 0; i < toMail.Count(); i++)
                {
                    Recipient toRecipient = new Recipient();
                    EmailAddress toEmailAddress = new EmailAddress();

                    toEmailAddress.Address = toMail[i];
                    toRecipient.EmailAddress = toEmailAddress;
                    toRecipients.Add(toRecipient);
                }

                List<Recipient> ccRecipients = new List<Recipient>();
                if (!string.IsNullOrEmpty(cc))
                {
                    string[] ccMail = cc.Split(';');                    
                    for (int j = 0; j < ccMail.Count(); j++)
                    {
                        Recipient ccRecipient = new Recipient();
                        EmailAddress ccEmailAddress = new EmailAddress();

                        ccEmailAddress.Address = ccMail[j];
                        ccRecipient.EmailAddress = ccEmailAddress;
                        ccRecipients.Add(ccRecipient);
                    }
                }

                List<string> lista_adjuntos = new List<string>();


                
                if (!string.IsNullOrEmpty(attachements))
                {
                    string[] adjunto = attachements.Split(';');
                    for (int j = 0; j < adjunto.Count(); j++)
                        lista_adjuntos.Add(adjunto[j]);

                }


                var attachments = new MessageAttachmentsCollectionPage();
                for (int j = 0; j < lista_adjuntos.Count(); j++)
                {
                    FileInfo archivo = new FileInfo(lista_adjuntos[j]);
                    FileAttachment file = new FileAttachment();

                    byte[] fileArray = System.IO.File.ReadAllBytes(archivo.FullName);

                    file.Name = archivo.Name;
                    file.ContentBytes = fileArray;
                    file.ContentId = archivo.Name;

                    attachments.Add(file);
                }


                    /*
                    var attachments = new MessageAttachmentsCollectionPage()
                    {
                    new FileAttachment{
                        ContentType= "image/jpeg",
                        ContentBytes = imageArray,
                        ContentId = imageID,
                        Name= "test-image"
                    }
                    };

                    */

                    var mailMessage = new Microsoft.Graph.Message
                {
                    Subject = subject,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = body
                    },
                    ToRecipients = toRecipients,
                    CcRecipients = ccRecipients,                    
                    Attachments = attachments


                };

                ficheroLog.Add("from: " + from + " to: " + to + " subject: " + subject);

                graphServiceClient
                   .Users[from]
                    .SendMail(mailMessage, true)
                    .Request()
                    .PostAsync().Wait();

                return "Email successfully sent.";

            }
            catch (Exception ex)
            {
                string cadena_error = ex.Message;
                Console.WriteLine("SendMail: " + ex.Message);
                if (ex.InnerException != null)
                {
                    Console.WriteLine("SendMail: " + ex.InnerException.Message);
                    ficheroLog.AddError("SendMail: " + ex.Message);
                    cadena_error += Environment.NewLine + ex.InnerException.Message;
                    if (ex.InnerException.InnerException != null)
                    {
                        cadena_error += Environment.NewLine + ex.InnerException.InnerException.Message;
                        ficheroLog.AddError("SendMail: " + ex.InnerException.InnerException.Message);
                    }
                }

                return cadena_error;
            }
        }

        public string SendMail(string from, string to, string cc, string subject, string body, string attachements, bool isHighImportace)
        {
            try
            {


                string[] toMail = to.Split(';');
                List<Recipient> toRecipients = new List<Recipient>();
                for (int i = 0; i < toMail.Count(); i++)
                {
                    Recipient toRecipient = new Recipient();
                    EmailAddress toEmailAddress = new EmailAddress();

                    toEmailAddress.Address = toMail[i];
                    toRecipient.EmailAddress = toEmailAddress;
                    toRecipients.Add(toRecipient);
                }

                List<Recipient> ccRecipients = new List<Recipient>();
                if (!string.IsNullOrEmpty(cc))
                {
                    string[] ccMail = cc.Split(';');
                    for (int j = 0; j < ccMail.Count(); j++)
                    {
                        Recipient ccRecipient = new Recipient();
                        EmailAddress ccEmailAddress = new EmailAddress();

                        ccEmailAddress.Address = ccMail[j];
                        ccRecipient.EmailAddress = ccEmailAddress;
                        ccRecipients.Add(ccRecipient);
                    }
                }

                List<string> lista_adjuntos = new List<string>();



                if (!string.IsNullOrEmpty(attachements))
                {
                    string[] adjunto = attachements.Split(';');
                    for (int j = 0; j < adjunto.Count(); j++)
                        lista_adjuntos.Add(adjunto[j]);

                }


                var attachments = new MessageAttachmentsCollectionPage();
                for (int j = 0; j < lista_adjuntos.Count(); j++)
                {
                    FileInfo archivo = new FileInfo(lista_adjuntos[j]);
                    FileAttachment file = new FileAttachment();

                    byte[] fileArray = System.IO.File.ReadAllBytes(archivo.FullName);

                    file.Name = archivo.Name;
                    file.ContentBytes = fileArray;
                    file.ContentId = archivo.Name;

                    attachments.Add(file);
                }


                /*
                var attachments = new MessageAttachmentsCollectionPage()
                {
                new FileAttachment{
                    ContentType= "image/jpeg",
                    ContentBytes = imageArray,
                    ContentId = imageID,
                    Name= "test-image"
                }
                };

                */

                var mailMessage = new Microsoft.Graph.Message
                {
                    Subject = subject,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = body
                    },
                    ToRecipients = toRecipients,
                    CcRecipients = ccRecipients,
                    Attachments = attachments,
                    Importance = (isHighImportace?Importance.High:Importance.Normal)

                };

                ficheroLog.Add("from: " + from + " to: " + to + " subject: " + subject);

                graphServiceClient
                   .Users[from]
                    .SendMail(mailMessage, true)
                    .Request()
                    .PostAsync().Wait();

                return "Email successfully sent.";

            }
            catch (Exception ex)
            {
                string cadena_error = ex.Message;
                Console.WriteLine("SendMail: " + ex.Message);
                if (ex.InnerException != null)
                {
                    Console.WriteLine("SendMail: " + ex.InnerException.Message);
                    ficheroLog.AddError("SendMail: " + ex.Message);
                    cadena_error += Environment.NewLine + ex.InnerException.Message;
                    if (ex.InnerException.InnerException != null)
                    {
                        cadena_error += Environment.NewLine + ex.InnerException.InnerException.Message;
                        ficheroLog.AddError("SendMail: " + ex.InnerException.InnerException.Message);
                    }
                }

                return cadena_error;
            }
        }
        public string SendMailNoHTML(string from, string to, string cc, string subject, string body, string attachements)
        {
            try
            {


                string[] toMail = to.Split(';');
                List<Recipient> toRecipients = new List<Recipient>();
                for (int i = 0; i < toMail.Count(); i++)
                {
                    Recipient toRecipient = new Recipient();
                    EmailAddress toEmailAddress = new EmailAddress();

                    toEmailAddress.Address = toMail[i];
                    toRecipient.EmailAddress = toEmailAddress;
                    toRecipients.Add(toRecipient);
                }

                List<Recipient> ccRecipients = new List<Recipient>();
                if (!string.IsNullOrEmpty(cc))
                {
                    string[] ccMail = cc.Split(';');
                    for (int j = 0; j < ccMail.Count(); j++)
                    {
                        Recipient ccRecipient = new Recipient();
                        EmailAddress ccEmailAddress = new EmailAddress();

                        ccEmailAddress.Address = ccMail[j];
                        ccRecipient.EmailAddress = ccEmailAddress;
                        ccRecipients.Add(ccRecipient);
                    }
                }

                List<string> lista_adjuntos = new List<string>();



                if (!string.IsNullOrEmpty(attachements))
                {
                    string[] adjunto = attachements.Split(';');
                    for (int j = 0; j < adjunto.Count(); j++)
                        lista_adjuntos.Add(adjunto[j]);

                }


                var attachments = new MessageAttachmentsCollectionPage();
                for (int j = 0; j < lista_adjuntos.Count(); j++)
                {
                    FileInfo archivo = new FileInfo(lista_adjuntos[j]);
                    FileAttachment file = new FileAttachment();

                    byte[] fileArray = System.IO.File.ReadAllBytes(archivo.FullName);

                    file.Name = archivo.Name;
                    file.ContentBytes = fileArray;
                    file.ContentId = archivo.Name;
                    attachments.Add(file);
                }


                /*
                var attachments = new MessageAttachmentsCollectionPage()
                {
                new FileAttachment{
                    ContentType= "image/jpeg",
                    ContentBytes = imageArray,
                    ContentId = imageID,
                    Name= "test-image"
                }
                };

                */




                var mailMessage = new Microsoft.Graph.Message
                {
                    Subject = subject,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = body
                    },
                    ToRecipients = toRecipients,
                    CcRecipients = ccRecipients,
                    Attachments = attachments
                    


                };

                

                ficheroLog.Add("from: " + from + " to: " + to + " subject: " + subject);

               

                graphServiceClient
                   .Users[from]
                    .SendMail(mailMessage, true)
                    .Request()
                    .PostAsync().Wait();

                return "Email successfully sent.";

            }
            catch (Exception ex)
            {
                string cadena_error = ex.Message;
                Console.WriteLine("SendMail: " + ex.Message);
                if (ex.InnerException != null)
                {
                    Console.WriteLine("SendMail: " + ex.InnerException.Message);
                    ficheroLog.AddError("SendMail: " + ex.Message);
                    cadena_error += Environment.NewLine + ex.InnerException.Message;
                    if (ex.InnerException.InnerException != null)
                    {
                        cadena_error += Environment.NewLine + ex.InnerException.InnerException.Message;
                        ficheroLog.AddError("SendMail: " + ex.InnerException.InnerException.Message);
                    }
                }

                return cadena_error;
            }
        }


        public string SaveMail(string from, string to, string cc, string subject, string body)
        {
            try
            {


                string[] toMail = to.Split(';');
                List<Recipient> toRecipients = new List<Recipient>();
                int i = 0;
                for (i = 0; i < toMail.Count(); i++)
                {
                    Recipient toRecipient = new Recipient();
                    EmailAddress toEmailAddress = new EmailAddress();

                    toEmailAddress.Address = toMail[i];
                    toRecipient.EmailAddress = toEmailAddress;
                    toRecipients.Add(toRecipient);
                }

                List<Recipient> ccRecipients = new List<Recipient>();
                if (!string.IsNullOrEmpty(cc))
                {
                    string[] ccMail = cc.Split(';');
                    int j = 0;
                    for (j = 0; j < ccMail.Count(); j++)
                    {
                        Recipient ccRecipient = new Recipient();
                        EmailAddress ccEmailAddress = new EmailAddress();

                        ccEmailAddress.Address = ccMail[j];
                        ccRecipient.EmailAddress = ccEmailAddress;
                        ccRecipients.Add(ccRecipient);
                    }
                }





                var mailMessage = new Microsoft.Graph.Message
                {
                    Subject = subject,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = body
                    },
                    ToRecipients = toRecipients,
                    CcRecipients = ccRecipients


                };

                ficheroLog.Add("from: " + from + " to: " + to + " subject: " + subject);

                graphServiceClient
                   .Users[from]
                    .SendMail(mailMessage, true)
                    .Request()
                    .PostAsync().Wait();

                return "Email successfully sent.";

            }
            catch (Exception ex)
            {
                string cadena_error = ex.Message;
                if (ex.InnerException != null)
                {
                    Console.WriteLine("SendMail: " + ex.InnerException.Message);
                    ficheroLog.AddError("SendMail: " + ex.Message);
                    cadena_error += Environment.NewLine + ex.InnerException.Message;
                    if (ex.InnerException.InnerException != null)
                    {
                        cadena_error += Environment.NewLine + ex.InnerException.InnerException.Message;
                        ficheroLog.AddError("SendMail: " + ex.InnerException.InnerException.Message);
                    }
                }

                return cadena_error;
            }
        }

        public string SaveMail(string from, string to, string cc, string subject, string body, string attachements)
        {
            try
            {


                string[] toMail = to.Split(';');
                List<Recipient> toRecipients = new List<Recipient>();
                for (int i = 0; i < toMail.Count(); i++)
                {
                    Recipient toRecipient = new Recipient();
                    EmailAddress toEmailAddress = new EmailAddress();

                    toEmailAddress.Address = toMail[i];
                    toRecipient.EmailAddress = toEmailAddress;
                    toRecipients.Add(toRecipient);
                }

                List<Recipient> ccRecipients = new List<Recipient>();
                if (!string.IsNullOrEmpty(cc))
                {
                    string[] ccMail = cc.Split(';');
                    for (int j = 0; j < ccMail.Count(); j++)
                    {
                        Recipient ccRecipient = new Recipient();
                        EmailAddress ccEmailAddress = new EmailAddress();

                        ccEmailAddress.Address = ccMail[j];
                        ccRecipient.EmailAddress = ccEmailAddress;
                        ccRecipients.Add(ccRecipient);
                    }
                }

                List<string> lista_adjuntos = new List<string>();



                if (!string.IsNullOrEmpty(attachements))
                {
                    string[] adjunto = attachements.Split(';');
                    for (int j = 0; j < adjunto.Count(); j++)
                        lista_adjuntos.Add(adjunto[j]);

                }


                var attachments = new MessageAttachmentsCollectionPage();
                for (int j = 0; j < lista_adjuntos.Count(); j++)
                {
                    FileInfo archivo = new FileInfo(lista_adjuntos[j]);
                    FileAttachment file = new FileAttachment();

                    byte[] fileArray = System.IO.File.ReadAllBytes(archivo.FullName);

                    file.Name = archivo.Name;
                    file.ContentBytes = fileArray;
                    file.ContentId = archivo.Name;

                    attachments.Add(file);
                }


                /*
                var attachments = new MessageAttachmentsCollectionPage()
                {
                new FileAttachment{
                    ContentType= "image/jpeg",
                    ContentBytes = imageArray,
                    ContentId = imageID,
                    Name= "test-image"
                }
                };

                */

                var mailMessage = new Microsoft.Graph.Message
                {
                    Subject = subject,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = body
                    },
                    ToRecipients = toRecipients,
                    CcRecipients = ccRecipients,
                    Attachments = attachments


                };

                ficheroLog.Add("from: " + from + " to: " + to + " subject: " + subject);

                graphServiceClient
                   .Users[from]
                    .SendMail(mailMessage, true)
                    .Request()
                    .PostAsync().Wait();

                return "Email successfully sent.";

            }
            catch (Exception ex)
            {
                string cadena_error = ex.Message;
                if (ex.InnerException != null)
                {
                    Console.WriteLine("SendMail: " + ex.InnerException.Message);
                    ficheroLog.AddError("SendMail: " + ex.Message);
                    cadena_error += Environment.NewLine + ex.InnerException.Message;
                    if (ex.InnerException.InnerException != null)
                    {
                        cadena_error += Environment.NewLine + ex.InnerException.InnerException.Message;
                        ficheroLog.AddError("SendMail: " + ex.InnerException.InnerException.Message);
                    }
                }

                return cadena_error;
            }
        }
    }
}
