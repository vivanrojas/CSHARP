using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.utilidades
{
    public class ListasCorreo
    {

        List<EndesaEntity.MailDetail> mailDetail = new List<EndesaEntity.MailDetail>();
        public EndesaEntity.Mail correo;

        public ListasCorreo(servidores.MySQLDB.Esquemas esquema, string tabla, string nombre_mail)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            correo = new EndesaEntity.Mail();

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
                    correo.mail_id = Convert.ToInt32(reader["mail_id"]);
                    correo.subject = reader["subject"].ToString();
                    correo.body = reader["body"].ToString();
                    correo.mailbox = reader["mailbox"] != System.DBNull.Value ? reader["mailbox"].ToString() : null;
                    if (Convert.ToInt32(DateTime.Now.ToString("HHmm")) >= 1400)
                        correo.body = correo.body.Replace("@1", "Buenas tardes");
                    else
                        correo.body = correo.body.Replace("@1", "Buenos días");


                    GetMailDetail(esquema, tabla, correo.mail_id);
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


                CreaDirecciones();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }


        private void CreaDirecciones()
        {
            bool firstOnly = true;
            try
            {

                for (int i = 0; i < mailDetail.Count; i++)
                {
                    if (mailDetail[i].mail_type == "to")
                    {
                        if (firstOnly)
                        {
                            correo.to = mailDetail[i].email;
                            firstOnly = false;
                        }
                            
                        else
                            correo.to = correo.to + ";" + mailDetail[i].email;
                    }                       
                    

                }



                firstOnly = true;
                for (int i = 0; i < mailDetail.Count; i++)
                {
                    if (mailDetail[i].mail_type == "cc")
                        if (firstOnly)
                        {
                            correo.cc = mailDetail[i].email;
                            firstOnly = false;
                        }                            
                        else
                            correo.cc = correo.cc + ";" + mailDetail[i].email;
                    

                }
                firstOnly = true;
                for (int i = 0; i < mailDetail.Count; i++)
                {
                    if (mailDetail[i].mail_type == "cco")
                        if (firstOnly)
                        {
                            correo.bcc = mailDetail[i].email;
                            firstOnly = false;
                        }
                            
                        else
                            correo.bcc = correo.bcc + ";" + mailDetail[i].email;
                    

                }

                

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }
    }
}
