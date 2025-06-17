using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion
{
    public class Agora_Portugal
    {

        Dictionary<string, string> dic_cups20;
        Dictionary<string, string> dic_cups13;
        public Agora_Portugal()
        {
            dic_cups13 = GetAll();
        }


        private Dictionary<string, string> GetAll()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string cups20 = "";
            string cups13 = "";
            Dictionary<string, string> d = new Dictionary<string, string>();

            try
            {
                strSql = "select cups13, cups20 from agora_portugal";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if(r["cups20"] != System.DBNull.Value)
                        cups20 = r["cups20"].ToString();

                    if (r["cups13"] != System.DBNull.Value)
                        cups13 = r["cups13"].ToString();

                    string o;
                    if (!d.TryGetValue(cups13, out o))
                        d.Add(cups13, cups13);
                }
                db.CloseConnection();

                strSql = "select CD_CUPS, CD_CUPS_EXT from agora_portugal_aqua";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["CD_CUPS"] != System.DBNull.Value)
                        cups13 = r["CD_CUPS"].ToString();

                    string o;
                    if (!d.TryGetValue(cups13, out o))
                        d.Add(cups13, cups13);
                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public bool EsAgora(string cups13)
        {
            string o;
            return dic_cups13.TryGetValue(cups13, out o);
        }






    }
}
