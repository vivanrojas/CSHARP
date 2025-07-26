using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.utilidades
{
    class GOMail
    {
        public enum destinatarios
        {
            to, cc, cco
        }

        public static List<string> Destinatarios(string proceso, destinatarios des)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string tipoDestinatario = "";
            List<string> lista = new List<string>();
            try
            {
                switch (des)
                {
                    case destinatarios.cc:
                        tipoDestinatario = "cc";
                        break;
                    case destinatarios.cco:
                        tipoDestinatario = "cco";
                        break;
                    case destinatarios.to:
                        tipoDestinatario = "to";
                        break;
                }

                strSql = "select mail from go_mail where"
                    + " process = '" + proceso + "' and"
                    + " sendertype = '" + tipoDestinatario + "';";

                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    lista.Add(reader["mail"].ToString());
                }
            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);
            }

            return lista;
        }

        public static string Destinatarios_lista(string proceso, destinatarios des)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string tipoDestinatario = "";
            string lista = "";
            bool firstOnly = true;

            try
            {
                switch (des)
                {
                    case destinatarios.cc:
                        tipoDestinatario = "cc";
                        break;
                    case destinatarios.cco:
                        tipoDestinatario = "cco";
                        break;
                    case destinatarios.to:
                        tipoDestinatario = "to";
                        break;
                }

                strSql = "select mail from go_mail where"
                    + " process = '" + proceso + "' and"
                    + " sendertype = '" + tipoDestinatario + "';";

                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    if (firstOnly)
                    {
                        lista = reader["mail"].ToString();
                        firstOnly = false;
                    }
                    else
                        lista = lista + ";" + reader["mail"].ToString();
                }
            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);
            }

            return lista;
        }
    }
}
