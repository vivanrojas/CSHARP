using EndesaBusiness.servidores;
using Microsoft.Exchange.WebServices.Data;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EndesaBusiness.office
{
    public class MailCompose : EndesaEntity.Mail
    {
        List<EndesaEntity.MailDetail> mailDetail = new List<EndesaEntity.MailDetail>();
        Outlook.Application myApp;
        Outlook.MailItem oMsg;
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;

        public string sender { get; set; }
        public List<string> para { get; set; }
        public List<string> cc { get; set; }
        public List<string> cco { get; set; }
        public string asunto { get; set; }
        public string htmlCuerpo { get; set; }
        public List<string> adjuntos { get; set; }
        public string mailbox { get; set; }


        public MailCompose()
        {


            para = new List<string>();
            cc = new List<string>();
            cco = new List<string>();
            adjuntos = new List<string>();

            myApp = new Outlook.Application();
            outlookNameSpace = myApp.GetNamespace("MAPI");

            //GetMail(esquema, tabla, nombre_mail);


            if (mailbox != null)
                inbox = outlookNameSpace.Folders[mailbox];
            else
                inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            oMsg = inbox.Items.Add();

        }

        public MailCompose(servidores.MySQLDB.Esquemas esquema, string tabla, string nombre_mail)
        {

            myApp = new Outlook.Application();
            outlookNameSpace = myApp.GetNamespace("MAPI");

            GetMail(esquema, tabla, nombre_mail);

            if (mailbox != null)
                inbox = outlookNameSpace.Folders[mailbox];
            else
                inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            oMsg = inbox.Items.Add();

        }



        //public void Send2()
        //{
        //    bool firstOnly = true;
        //    try
        //    {
        //        for (int i = 0; i < para.Count(); i++)
        //        {
        //            if (firstOnly)
        //                oMsg.To = para[i];
        //            else
        //                oMsg.To = oMsg.To + ";" + para[i];
        //        }

        //        firstOnly = true;
        //        for (int i = 0; i < cc.Count(); i++)
        //        {
        //            if (firstOnly)
        //                oMsg.CC = cc[i];
        //            else
        //                oMsg.CC = oMsg.CC + ";" + cc[i];
        //        }

        //        firstOnly = true;
        //        for (int i = 0; i < cco.Count(); i++)
        //        {
        //            if (firstOnly)
        //                oMsg.BCC = cco[i];
        //            else
        //                oMsg.BCC = oMsg.BCC + ";" + cco[i];
        //        }

        //        firstOnly = true;
        //        for (int i = 0; i < adjuntos.Count(); i++)
        //        {
        //            oMsg.Attachments.Add(adjuntos[i]);
        //        }

        //        oMsg.Subject = asunto;
        //        if (htmlCuerpo.Contains("HTML"))
        //            oMsg.HTMLBody = htmlCuerpo;
        //        else
        //            oMsg.Body = htmlCuerpo;

        //        oMsg.Send();


        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.Message, "MailCompose.Send",
        //         MessageBoxButtons.OK,
        //         MessageBoxIcon.Error);
        //    }

        //}


        public void Send()
        {

            string _to = "";
            string _cc = "";
            string _bcc = "";

            bool firstOnly = true;
            try
            {

                for (int i = 0; i < mailDetail.Count; i++)
                {
                    if (mailDetail[i].mail_type == "to")
                        if (firstOnly)
                            oMsg.To = mailDetail[i].email;
                        else
                            oMsg.To = oMsg.To + ";" + mailDetail[i].email;
                    firstOnly = false;

                }
                firstOnly = true;
                for (int i = 0; i < mailDetail.Count; i++)
                {
                    if (mailDetail[i].mail_type == "cc")
                        if (firstOnly)
                            oMsg.CC = mailDetail[i].email;
                        else
                            oMsg.CC = oMsg.CC + ";" + mailDetail[i].email;
                    firstOnly = false;

                }
                firstOnly = true;
                for (int i = 0; i < mailDetail.Count; i++)
                {
                    if (mailDetail[i].mail_type == "cco")
                        if (firstOnly)
                            oMsg.BCC = mailDetail[i].email;
                        else
                            oMsg.BCC = oMsg.BCC + ";" + mailDetail[i].email;
                    firstOnly = false;

                }

                oMsg.Subject = subject;
                oMsg.Body = body;
                oMsg.Send();

                //EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
                //mes.SendMail(mailbox, to, null, subject, body, null);

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        private void GetMail(servidores.MySQLDB.Esquemas esquema, string tabla, string nombre_mail)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            try
            {
                strSql = "select * from " + tabla + " where"
                    + " name = '" + nombre_mail + "' and "
                    + " (begin_date <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' and"
                    + " end_date >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "')";
                db = new MySQLDB(esquema);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    mail_id = Convert.ToInt32(reader["mail_id"]);
                    subject = reader["subject"].ToString();
                    body = reader["body"].ToString();
                    mailbox = reader["mailbox"] != System.DBNull.Value ? reader["mailbox"].ToString() : null;
                    if (Convert.ToInt32(DateTime.Now.ToString("HHmm")) >= 1400)
                        body = body.Replace("@1", "Buenas tardes");
                    else
                        body = body.Replace("@1", "Buenos días");


                    GetMailDetail(esquema, tabla, mail_id);
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void GetMailDetail(servidores.MySQLDB.Esquemas esquema, string tabla, int mail_id)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            try
            {
                strSql = "select * from " + tabla + "_detail" + " where"
                    + " mail_id = " + mail_id;

                db = new MySQLDB(esquema);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    EndesaEntity.MailDetail md = new EndesaEntity.MailDetail();
                    md.mail_id = mail_id;
                    md.email = reader["mail"].ToString();
                    md.mail_type = reader["mail_type"].ToString();
                    mailDetail.Add(md);

                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
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
                MessageBox.Show(e.Message, "MailCompose.Send",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
            }

        }
    }
}
