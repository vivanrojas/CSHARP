using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion.eexxi
{
    class Enviar_Mails
    {
        public Dictionary<string, Mail_Plantilla> dic_plantillas { get; set; }
        public Dictionary<string, List<Mail_Destinatarios>> dic_destinatarios { get; set; }
        public Enviar_Mails()
        {

            dic_plantillas = CargaPlantillasMail();
            dic_destinatarios = CargaDestinatarios();
        }

        private Dictionary<string, List<Mail_Destinatarios>> CargaDestinatarios()
        {
            Dictionary<string, List<Mail_Destinatarios>> dic = new Dictionary<string, List<Mail_Destinatarios>>();
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string plantilla = "";

            strSql = "select nombre_plantilla, mail, tipo_destinatario from eexxi_mail_destinatarios";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                Mail_Destinatarios c = new Mail_Destinatarios();
                if (r["nombre_plantilla"] != System.DBNull.Value)
                    plantilla = r["nombre_plantilla"].ToString();
                if (r["mail"] != System.DBNull.Value)
                    c.direccion_mail = r["mail"].ToString();
                if (r["tipo_destinatario"] != System.DBNull.Value)
                    c.tipo_destinatario = r["tipo_destinatario"].ToString();

                List<Mail_Destinatarios> o;
                if (!dic.TryGetValue(plantilla, out o))
                {
                    o = new List<Mail_Destinatarios>();
                    o.Add(c);
                    dic.Add(plantilla, o);
                }
                else
                    o.Add(c);

            }
            db.CloseConnection();

            return dic;
        }

        private Dictionary<string, Mail_Plantilla> CargaPlantillasMail()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            Dictionary<string, Mail_Plantilla> dic = new Dictionary<string, Mail_Plantilla>();

            strSql = "select nombre_plantilla, asunto, cuerpo from eexxi_mail_plantillas";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                Mail_Plantilla c = new Mail_Plantilla();

                if (r["nombre_plantilla"] != System.DBNull.Value)
                    c.nombre_plantilla = r["nombre_plantilla"].ToString();

                if (r["asunto"] != System.DBNull.Value)
                    c.asunto = r["asunto"].ToString();

                if (r["cuerpo"] != System.DBNull.Value)
                    c.cuerpo_mail = r["cuerpo"].ToString();

                Mail_Plantilla o;
                if (!dic.TryGetValue(c.nombre_plantilla, out o))
                    dic.Add(c.nombre_plantilla, c);
            }
            db.CloseConnection();

            return dic;
        }
    }
}
