using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EndesaBusiness.office
{
    public class SendMail
    {
        // Outlook.NameSpace outlookNameSpace;
        // Outlook.MAPIFolder inbox;
        //  Outlook.Items items;
        Outlook.Application myApp;
        Outlook.MailItem oMsg;
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;

        public string sender { get; set; }
        public List<string> para { get; set; }
        public List<string> cc { get; set; }
        public List<string> cco { get; set; }
        public string adjunto { get; set; }
        public string asunto { get; set; }
        public string htmlCuerpo { get; set; }
        public List<string> adjuntos { get; set; }



        public SendMail()
        {
            para = new List<string>();
            cc = new List<string>();
            cco = new List<string>();
            adjuntos = new List<string>();

            myApp = new Outlook.Application();
            oMsg = (Outlook.MailItem)myApp.CreateItem(Outlook.OlItemType.olMailItem);
        }

        public SendMail(string fromMailBox)
        {
            para = new List<string>();
            cc = new List<string>();
            cco = new List<string>();
            adjuntos = new List<string>();

            myApp = new Outlook.Application();
            outlookNameSpace = myApp.GetNamespace("MAPI");

            //GetMail(esquema, tabla, nombre_mail);

            if (fromMailBox != null)
                inbox = outlookNameSpace.Folders[fromMailBox];
            else
                inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            oMsg = inbox.Items.Add();
        }


        public void Send()
        {
            bool firstOnly = true;
            Outlook.Account account;
            
            try
            {

                for (int i = 0; i < para.Count; i++)
                {
                    if (firstOnly)
                    {
                        oMsg.To = para[i];
                        firstOnly = false;
                    }
                    else
                    {
                        oMsg.To = oMsg.To + ";" + para[i];
                    }

                }
                firstOnly = true;
                for (int i = 0; i < cc.Count; i++)
                {
                    if (firstOnly)
                    {
                        oMsg.CC = cc[i];
                        firstOnly = false;
                    }
                    else
                    {
                        oMsg.CC = oMsg.CC + ";" + cc[i];
                    }
                }
                firstOnly = true;
                for (int i = 0; i < cco.Count; i++)
                {
                    if (firstOnly)
                    {
                        oMsg.BCC = cco[i];
                        firstOnly = false;
                    }
                    else
                    {
                        oMsg.BCC = oMsg.BCC + ";" + cco[i];
                    }

                }

                firstOnly = true;
                for (int i = 0; i < adjuntos.Count; i++)
                {
                    oMsg.Attachments.Add(adjuntos[i]);
                }
                               

                oMsg.Subject = asunto;

                if (htmlCuerpo.Contains("HTML"))
                    oMsg.HTMLBody = htmlCuerpo;
                else
                    oMsg.Body = htmlCuerpo;

                if (sender != null)
                {
                    account = GetAccountForEmailAddress(myApp, sender);
                    oMsg.SendUsingAccount = account;
                }
                //oMsg.Subject = asunto;
                //oMsg.Body = htmlCuerpo;

                //if (adjunto != null)
                //{

                //    Outlook.Attachments attachments = oMsg.Attachments;
                //    attachments.Add(adjunto);
                //}

                oMsg.Send();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }
        public void Save()
        {
            bool firstOnly = true;
            try
            {
                for (int i = 0; i < para.Count; i++)
                {
                    if (firstOnly)
                    {
                        oMsg.To = para[i];
                        firstOnly = false;
                    }
                    else
                    {
                        oMsg.To = oMsg.To + ";" + para[i];
                    }

                }
                firstOnly = true;
                for (int i = 0; i < cc.Count; i++)
                {
                    if (firstOnly)
                    {
                        oMsg.CC = cc[i];
                        firstOnly = false;
                    }
                    else
                    {
                        oMsg.CC = oMsg.CC + ";" + cc[i];
                    }
                }
                firstOnly = true;
                for (int i = 0; i < cco.Count; i++)
                {
                    if (firstOnly)
                    {
                        oMsg.BCC = cco[i];
                        firstOnly = false;
                    }
                    else
                    {
                        oMsg.BCC = oMsg.BCC + ";" + cco[i];
                    }

                }


                firstOnly = true;
                for (int i = 0; i < adjuntos.Count; i++)
                {
                    oMsg.Attachments.Add(adjuntos[i]);
                }

                oMsg.Subject = asunto;
                
                if (htmlCuerpo.Contains("HTML"))
                    oMsg.HTMLBody = htmlCuerpo;
                else
                    oMsg.Body = htmlCuerpo;

                if (adjunto != null)
                {

                    Outlook.Attachments attachments = oMsg.Attachments;
                    attachments.Add(adjunto);
                }

                oMsg.Save();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        private void OMsg_AttachmentAdd1(Outlook.Attachment Attachment)
        {
            throw new NotImplementedException();
        }

        private void OMsg_AttachmentAdd(Outlook.Attachment Attachment)
        {
            throw new NotImplementedException();
        }

        public Outlook.Account GetAccountForEmailAddress(Outlook.Application application, string smtpAddress)
        {

            // Loop over the Accounts collection of the current Outlook session. 
            Outlook.Accounts accounts = application.Session.Accounts;
            foreach (Outlook.Account account in accounts)
            {
                // When the e-mail address matches, return the account. 
                if (account.SmtpAddress == smtpAddress)
                {
                    return account;
                }
            }
            throw new System.Exception(string.Format("No existe ninguna cuenta con la dirección: {0}", smtpAddress));
        }
    }
}
