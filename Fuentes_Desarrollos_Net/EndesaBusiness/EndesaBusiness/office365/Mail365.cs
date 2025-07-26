using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Xml;
using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System.Threading;
using System.Net;

namespace EndesaBusiness.office365
{
    public class Mail365
    {
        ExchangeService service;
        Mailbox mb;
        FolderId fid;
        FolderId outboxNoNuestros;
        FolderId outboxNuestros;
        Folder inbox;
        Folder outbox;
        EndesaEntity.ReglaCorreo regla;

        utilidades.Param param;
        utilidades.ParamUser paramUser;

        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "office365");


        public string body = "";
        public string subject = "";
        public string from = "";
        public string to = "";
        public string cc = null;
        public string adjuntos = null;

        utilidades.Credenciales credenciales;

        public Mail365(string user, string buzon)
        {
            credenciales = new utilidades.Credenciales(user);
            
            service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);            
            service.Credentials = new WebCredentials("rsiope.gma@enel.com", credenciales.server_password);
            service.AutodiscoverUrl(credenciales.server_user, RedirectionUrlValidationCallback);
            //ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
            
            mb = new Mailbox("rsiope.gma@enel.com");
            fid = new FolderId(WellKnownFolderName.Inbox, mb);            
            inbox = Folder.Bind(service, fid);

        }


        public void BorrarPapelera()
        {
            if (inbox != null)
            {
                FindItemsResults<Item> items = inbox.FindItems(new ItemView(300));

                foreach (EmailMessage eMail in items)
                {
                    eMail.Delete(DeleteMode.HardDelete);
                }
            }
        }


        

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 

            ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;


            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        private void GetAttachmentsFromEmail(ExchangeService service, ItemId itemId)
        {
            // Bind to an existing message item and retrieve the attachments collection.
            // This method results in an GetItem call to EWS.
            EmailMessage message = EmailMessage.Bind(service, itemId, new PropertySet(ItemSchema.Attachments));

            // Iterate through the attachments collection and load each attachment.
            foreach (Attachment attachment in message.Attachments)
            {
                if (attachment is FileAttachment)
                {
                    FileAttachment fileAttachment = attachment as FileAttachment;

                    // Load the attachment into a file.
                    // This call results in a GetAttachment call to EWS.
                    fileAttachment.Load("C:\\temp\\pnts\\" + fileAttachment.Name);

                    Console.WriteLine("File attachment name: " + fileAttachment.Name);
                }
                else // Attachment is an item attachment.
                {
                    ItemAttachment itemAttachment = attachment as ItemAttachment;

                    // Load attachment into memory and write out the subject.
                    // This does not save the file like it does with a file attachment.
                    // This call results in a GetAttachment call to EWS.
                    itemAttachment.Load();

                    Console.WriteLine("Item attachment name: " + itemAttachment.Name);
                }
            }
        }
                
        public void GeneraNuevoCorreo()
        {

           
            string[] listaArchivos;
            FileInfo file;
            bool firstOnly = true;

           






            //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer(System.Environment.UserName);
            //foreach (Attachment attachment in mail.Attachments)
            //{


            //    string filePath = Path.Combine(regla.rutaSalvadoAdjuntos, attachment.Name);
            //    FileAttachment fileAttachment = attachment as FileAttachment;
            //    fileAttachment.Load(filePath);
            //    if (firstOnly)
            //    {
            //        adjuntos = filePath;
            //        firstOnly = false;
            //    }

            //    else
            //        adjuntos += ";" + filePath;


            //}

            EmailMessage email = new EmailMessage(service);
            email.Body = new MessageBody(BodyType.HTML, body.ToString());
            email.From = new EmailAddress(from);
            
            Mailbox SharedMailbox = new Mailbox(credenciales.mail);
            FolderId SharedMailboxSendItems = new FolderId(WellKnownFolderName.SentItems, SharedMailbox);
            email.SendAndSaveCopy(SharedMailboxSendItems);


            //mes.SendMailWeb(from, to, cc, subject, body, adjuntos);
            //Thread.Sleep(1000);
            //mes = null;

            /*
           
            */

        }

        

        private FolderId GetFolderID(string folderName)
        {
            FolderId fid = new FolderId(WellKnownFolderName.MsgFolderRoot, mb);
            Folder rootfolder = Folder.Bind(service, fid);
            rootfolder.Load();

            foreach (Folder folder in rootfolder.FindFolders(new FolderView(100)))
            {
                // Finds the emails in a certain folder, in this case the Junk Email

                // Console.WriteLine(folder.DisplayName);
                // This IF limits what folder the program will seek
                if (folder.DisplayName == folderName)
                {
                    // Trust me, the ID is a pain if you want to manually copy and paste it. This stores it in a variable
                    return folder.Id;                  
                    
                }
            }

            return null;
        }

    }
}
