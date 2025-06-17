using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EndesaBusiness.office
{
    class CreaCorreo
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "GBL_CreaCorreo");

        public string sender { get; set; }
        public List<string> para { get; set; }
        public List<string> cc { get; set; }
        public List<string> cco { get; set; }
        public string asunto { get; set; }
        public string htmlCuerpo { get; set; }
        public string cuerpoWeb { get; set; }
        public List<string> adjuntos { get; set; }
        public string mailbox { get; set; }
        public string subFolderMailBox { get; set; }


        // List<CorreoDetalle> mailDetail = new List<CorreoDetalle>();
        private Outlook.Application myApp;
        private Outlook.MailItem oMsg;
        private Outlook.NameSpace outlookNameSpace;
        private Outlook.MAPIFolder inbox;

        public CreaCorreo(string fromEmailAccount)
        {
            para = new List<string>();
            cc = new List<string>();
            cco = new List<string>();
            adjuntos = new List<string>();

            myApp = new Outlook.Application();
            outlookNameSpace = myApp.GetNamespace("MAPI");
            Outlook.Account account = GetAccountForEmailAddress(myApp, fromEmailAccount);


            // inbox = outlookNameSpace.Folders[mailbox].Folders[subFolderMailBox];
            inbox = outlookNameSpace.Folders[account.ExchangeMailboxServerName];
            //else
            //    inbox = outlookNameSpace.GetDefaultFolder(
            //        Microsoft.Office.Interop.Outlook.
            //        OlDefaultFolders.olFolderInbox);

            oMsg = inbox.Items.Add();
            //oMsg = new Outlook.MailItem();
        }


        public CreaCorreo()
        {

            para = new List<string>();
            cc = new List<string>();
            cco = new List<string>();
            adjuntos = new List<string>();

            myApp = new Outlook.Application();
            outlookNameSpace = myApp.GetNamespace("MAPI");

            // GetMail(esquema, tabla, nombre_mail);

        }

        public void Send()
        {

            if (mailbox != null)
                inbox = outlookNameSpace.Folders[mailbox];
            else
                inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            oMsg = inbox.Items.Add();

            bool firstOnly = true;
            try
            {
                for (int i = 0; i < para.Count(); i++)
                {
                    if (firstOnly)
                    {
                        oMsg.To = para[i];
                        firstOnly = false;
                    }

                    else
                        oMsg.To = oMsg.To + ";" + para[i];
                }

                firstOnly = true;
                for (int i = 0; i < cc.Count(); i++)
                {
                    if (firstOnly)
                    {
                        oMsg.CC = cc[i];
                        firstOnly = false;
                    }
                    else
                        oMsg.CC = oMsg.CC + ";" + cc[i];
                }

                firstOnly = true;
                for (int i = 0; i < cco.Count(); i++)
                {
                    if (firstOnly)
                    {
                        oMsg.BCC = cco[i];
                        firstOnly = false;
                    }
                    else
                        oMsg.BCC = oMsg.BCC + ";" + cco[i];
                }

                firstOnly = true;
                for (int i = 0; i < adjuntos.Count(); i++)
                {
                    oMsg.Attachments.Add(adjuntos[i]);
                }

                oMsg.Subject = asunto;
                if (htmlCuerpo != null)
                    oMsg.Body = htmlCuerpo;
                if (cuerpoWeb != null)
                    oMsg.HTMLBody = cuerpoWeb;

                oMsg.Send();

            }
            catch (Exception e)
            {
                ficheroLog.AddError(e.Message);
            }

        }
        public void Save()
        {

            bool firstOnly = true;
            try
            {
                if (mailbox != null)
                    inbox = outlookNameSpace.Folders[mailbox];
                else
                    inbox = outlookNameSpace.GetDefaultFolder(
                        Microsoft.Office.Interop.Outlook.
                        OlDefaultFolders.olFolderInbox);

                oMsg = inbox.Items.Add();


                for (int i = 0; i < para.Count(); i++)
                {
                    if (firstOnly)
                    {
                        oMsg.To = para[i];
                        firstOnly = false;
                    }

                    else
                        oMsg.To = oMsg.To + ";" + para[i];
                }

                firstOnly = true;
                for (int i = 0; i < cc.Count(); i++)
                {
                    if (firstOnly)
                    {
                        oMsg.CC = cc[i];
                        firstOnly = false;
                    }
                    else
                        oMsg.CC = oMsg.CC + ";" + cc[i];
                }

                firstOnly = true;
                for (int i = 0; i < cco.Count(); i++)
                {
                    if (firstOnly)
                    {
                        oMsg.BCC = cco[i];
                        firstOnly = false;
                    }

                    else
                        oMsg.BCC = oMsg.BCC + ";" + cco[i];
                }

                firstOnly = true;
                for (int i = 0; i < adjuntos.Count(); i++)
                {
                    oMsg.Attachments.Add(adjuntos[i]);
                }

                oMsg.Subject = asunto;
                if (htmlCuerpo != null)
                    oMsg.Body = htmlCuerpo;
                if (cuerpoWeb != null)
                    oMsg.HTMLBody = cuerpoWeb;

                oMsg.Save();

            }
            catch (Exception e)
            {
                ficheroLog.AddError("CreaCorreo: " + e.Message);
            }

        }

        public void Response()
        {
            oMsg.Reply();
        }

        public void Show()
        {



            bool firstOnly = true;
            try
            {
                for (int i = 0; i < para.Count(); i++)
                {
                    if (firstOnly)
                        oMsg.To = para[i];
                    else
                        oMsg.To = oMsg.To + ";" + para[i];
                }

                firstOnly = true;
                for (int i = 0; i < cc.Count(); i++)
                {
                    if (firstOnly)
                        oMsg.CC = cc[i];
                    else
                        oMsg.CC = oMsg.To + ";" + cc[i];
                }

                firstOnly = true;
                for (int i = 0; i < cco.Count(); i++)
                {
                    if (firstOnly)
                        oMsg.BCC = cco[i];
                    else
                        oMsg.BCC = oMsg.To + ";" + cco[i];
                }

                firstOnly = true;
                for (int i = 0; i < adjuntos.Count(); i++)
                {
                    oMsg.Attachments.Add(adjuntos[i]);
                }

                oMsg.Subject = asunto;
                oMsg.Body = htmlCuerpo;
                oMsg.Display();


            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        private Outlook.Account GetAccountForEmailAddress(Outlook.Application application, string smtpAddress)
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
            throw new System.Exception(string.Format("No Account with SmtpAddress: {0} exists!", smtpAddress));
        }

    }
}
