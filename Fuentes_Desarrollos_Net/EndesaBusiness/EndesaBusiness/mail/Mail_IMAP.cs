using System;
using System.Collections.Generic;
using System.Linq;

using System.Text;
using System.Threading.Tasks;
using IMAPMail = AE.Net.Mail;

namespace EndesaBusiness.mail
{
    public class Mail_IMAP
    {
        string _server;
        string _user;
        string _pass;

        public Mail_IMAP(string server, string user, string pass)
        {
            _server = server;
            _user = user;
            _pass = pass;
        }

        public void ReadEmailBox()
        {
            IMAPMail.ImapClient ic = new IMAPMail.ImapClient(_server, _user, _pass, IMAPMail.AuthMethods.Login, 993, true);
            ic.SelectMailbox("INBOX");
            Console.WriteLine(ic.GetMessageCount());
            IMAPMail.MailMessage[] mm = ic.GetMessages(0, 10);
            foreach (IMAPMail.MailMessage m in mm)
            {
                Console.WriteLine(m.Subject);
                
            }
            
            ic.Dispose();

        }
    }
}
