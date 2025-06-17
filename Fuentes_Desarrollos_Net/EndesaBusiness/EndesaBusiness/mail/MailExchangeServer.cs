using EndesaBusiness.servidores;
using Microsoft.Exchange.WebServices.Data;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using System.Text;
using System.Threading.Tasks;
using MailKit.Security;
using MailKit.Net.Pop3;
using System.Security.Authentication;
//using IMAPMail = AE.Net.Mail;



namespace EndesaBusiness.mail
{
    public class MailExchangeServer
    {
        private string _server;
        private string _user;
        private string _pass;
        private int _port;

        public MailExchangeServer(string server, string user, string pass, int port)
        {
            _server = server;
            _user = user;
            _pass = pass;
            _port = port;
        }

        public MailExchangeServer()
        {
            GetMailServerParameters(null);
        }

        public MailExchangeServer(string user)
        {
            GetMailServerParameters(user);
        }

        public void SendMail(string from, string to, string cc, string subject, string body, string attachment)
        {
            SmtpClient client = new SmtpClient();            
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential(_user, _pass);
            client.Port = _port;
            
            client.Host = _server;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.EnableSsl = true;
            
            //Posiblemente borrar
            //client.Port = 993;

            MailAddress _from = new MailAddress(from, String.Empty, System.Text.Encoding.UTF8);            
            string[] lista = to.Split(';');            

            
            MailMessage message = new MailMessage(from, lista[0]);
            for (int i = 1; i < lista.Count(); i++)
            {
                message.To.Add(lista[i]);
            }
           

            message.IsBodyHtml = false;

            if (cc != null)
            {
                string[] lista_cc = cc.Split(';');
                for (int i = 0; i < lista_cc.Count(); i++)
                {
                    MailAddress _cc = new MailAddress(lista_cc[i]);
                    message.CC.Add(_cc);
                }
                  
                
            }

            if(attachment != null)
            {
                message.Attachments.Add(new System.Net.Mail.Attachment(attachment));
            }

            message.Body = body;            
            message.Subject = subject;
            message.SubjectEncoding = System.Text.Encoding.UTF8;            

            client.Send(message);
            
        }

        public void SaveMail(string from, string to, string cc, string subject, string body, string attachment)
        {
            SmtpClient client = new SmtpClient();
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential(_user, _pass);
            client.Port = _port;
            
            client.Host = _server;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.EnableSsl = true;
            bool firstOnly = true;

            MailAddress _from = new MailAddress(from, String.Empty, System.Text.Encoding.UTF8);
            string[] lista = to.Split(';');


            MailMessage message = new MailMessage(from, lista[0]);
            for (int i = 1; i < lista.Count(); i++)
            {
                message.To.Add(lista[i]);
            }

            message.IsBodyHtml = false;

            if (cc != null)
            {
                string[] lista_cc = cc.Split(';');
                MailAddress _cc = new MailAddress(lista[0]);
                for (int i = 0; i < lista_cc.Count(); i++)
                {
                    MailAddress __cc = new MailAddress(lista_cc[i]);
                    message.CC.Add(__cc);
                }

            }

            if (attachment != null)
            {
                message.Attachments.Add(new System.Net.Mail.Attachment(attachment));
            }

            message.Body = body;            
            message.Subject = subject;
            message.SubjectEncoding = System.Text.Encoding.UTF8;

            client.PickupDirectoryLocation = @"c:\Temp\";

            

        }

        public void SendMailWeb(string from, string to, string cc, string subject, string body, string attachment)
        {
            SmtpClient client = new SmtpClient();
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential(_user, _pass);
            client.Port = _port;
            
            client.Host = _server;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.EnableSsl = true;
            

            MailAddress _from = new MailAddress(from, String.Empty, System.Text.Encoding.UTF8);
            string[] lista = to.Split(';');

            MailMessage message = new MailMessage(from, lista[0]);
            for (int i = 1; i < lista.Count(); i++)
            {
                message.To.Add(lista[i]);
            }

            message.IsBodyHtml = true;

            if (cc != null)
            {
                lista = cc.Split(';');
                MailAddress _cc = new MailAddress(lista[0]);
                for (int i = 0; i < lista.Count(); i++)
                {
                    message.CC.Add(lista[i]);
                }

            }

            if (attachment != null)
            {
                string[] lista_adjuntos = attachment.Split(';');
                if (lista_adjuntos.Count() > 0)
                {
                   for(int i = 0; i < lista_adjuntos.Count(); i++)     
                        message.Attachments.Add(new System.Net.Mail.Attachment(lista_adjuntos[i]));
                }
            }
                
            

            message.Body = body;            
            message.Subject = subject;
            message.SubjectEncoding = System.Text.Encoding.UTF8;
            client.Send(message);

            if (message.Attachments != null)
            {
                for(int i = message.Attachments.Count - 1; i >= 0; i--)
                {
                    message.Attachments[i].Dispose();
                }
                message.Attachments.Clear();
                message.Attachments.Dispose();

            }

            message.Dispose();
            message = null;
            client = null;
            
        }
               

        private void GetMailServerParameters(string user)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                if (user == null)
                    user = System.Environment.UserName.ToUpper();
                

                strSql = "SELECT user_id, user_name, from_date, to_date, mail, server, server_user, server_password, server_port"
                    + " FROM aux1.mail_users_loging_info where "
                    + " user_id = '" +  user.ToUpper() + "' and"
                    + " (from_date <= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' and" 
                    + " to_date >= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";                
                db = new MySQLDB(servidores.MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    _server = r["server"].ToString();
                    _user = r["server_user"].ToString();
                    _pass = utilidades.FuncionesTexto.Decrypt(r["server_password"].ToString(), true);
                    _port = Convert.ToInt32(r["server_port"]);
                }
                db.CloseConnection();
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        public void MoveMail()
        {
            
        }

        public void Old_ReadFolder()
        {
            //// Connect to the IMAP server. The 'true' parameter specifies to use SSL
            //// which is important (for Gmail at least)
            //IMAPMail.ImapClient ic = new IMAPMail.ImapClient("outlook.office365.com", "rsiope.gma@enel.com", "8!e3QMgUq11!",
            //                 IMAPMail.AuthMethods.Login, 993, true);
            //// Select a mailbox. Case-insensitive
            //// IMAPMail.Imap.Mailbox[] lista = ic.ListMailboxes(string.Empty,"*");
            //ic.Namespace();
            //ic.SelectMailbox("Endesa Energia, PNT");
            //// ic.SuscribeMailbox("eepnt@enel.com");
            //Console.WriteLine(ic.GetMessageCount());
            //// Get the first *11* messages. 0 is the first message;
            //// and it also includes the 10th message, which is really the eleventh ;)
            //// MailMessage represents, well, a message in your mailbox
            //IMAPMail.MailMessage[] mm = ic.GetMessages(0, 5);
            //foreach (IMAPMail.MailMessage m in mm)
            //{
            //    Console.WriteLine(m.Subject);
            //}

            //ic.Dispose();

            //// https://outlook.office365.com/EWS/Exchange.asmx

        }


        public void ReadFolderPOP()
        {
            using (var client = new Pop3Client(new ProtocolLogger("pop.log")))
            {
                client.Connect("outlook.office365.com", 995, SecureSocketOptions.SslOnConnect);                
                
                client.Authenticate("rsiope.gma@enel.com", "@@Peti100001@@");
                

                for (int i = 0; i < client.Count; i++)
                {
                    var message = client.GetMessage(i);
                    Console.WriteLine("Subject: {0}", message.Subject);
                }

                client.Disconnect(true);
            }
        }

        public void ReadFolder()
        {
            try
            {

                var client = new ImapClient(new ProtocolLogger("imap.log"));
                //client.SslProtocols = SslProtocols.Ssl3 | SslProtocols.Tls | SslProtocols.Tls11 | SslProtocols.Tls12;
                //client.CheckCertificateRevocation = false;
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                client.Connect("outlook.office365.com", 993, SecureSocketOptions.SslOnConnect);
                client.Authenticate("rsiope.gma@enel.com", "@@Peti100001@@");
                //client.AuthenticationMechanisms.Remove("XOAUTH2");

                var inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadOnly);
                Console.WriteLine("Total messages: {0}", inbox.Count);
                Console.WriteLine("Recent messages: {0}", inbox.Recent);

                for (int i = 0; i < inbox.Count; i++)
                {
                    var message = inbox.GetMessage(i);
                    Console.WriteLine("Subject: {0}", message.Subject);
                }

                client.Disconnect(true);               
                
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        private static SearchFilter SetFilter()
        {
            List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
            searchFilterCollection.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
            searchFilterCollection.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, true));
            SearchFilter s = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, searchFilterCollection.ToArray());
            return s;
        }

        private bool CertificateValidationCallBack(
             object sender,
             System.Security.Cryptography.X509Certificates.X509Certificate certificate,
             System.Security.Cryptography.X509Certificates.X509Chain chain,
             System.Net.Security.SslPolicyErrors sslPolicyErrors)
            {
                // If the certificate is a valid, signed certificate, return true.
                if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
                {
                    return true;
                }

                // If there are errors in the certificate chain, look at each error to determine the cause.
                if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
                {
                    if (chain != null && chain.ChainStatus != null)
                    {
                        foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                        {
                            if ((certificate.Subject == certificate.Issuer) &&
                               (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                            {
                                // Self-signed certificates with an untrusted root are valid. 
                                continue;
                            }
                            else
                            {
                                if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                                {
                                    // If there are any other errors in the certificate chain, the certificate is invalid,
                                    // so the method returns false.
                                    return false;
                                }
                            }
                        }
                    }

                    // When processing reaches this line, the only errors in the certificate chain are 
                    // untrusted root errors for self-signed certificates. These certificates are valid
                    // for default Exchange server installations, so return true.
                    return true;
                }
                else
                {
                    // In all other cases, return false.
                    return false;
                }
            }
        }
}
