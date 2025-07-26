using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion
{
    public class Agrupadas
    {

        // Dictionary<contraext, ccomtcom> dic;
        Dictionary<string, string> dic;

        public Agrupadas()
        {
            dic = Carga();
        }

        private Dictionary<string, string> Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            string contraext = "";
            string ccomtcom = "";

            Dictionary<string, string> d = new Dictionary<string, string>();
            try
            {                

                strSql = "select contraext, ccomtcom from irf_cont_fact_agrupada"
                    + " where tfacagru = 'S'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    contraext = r["contraext"].ToString();
                    ccomtcom = r["ccomtcom"].ToString(); 

                    string o;
                    if (!d.TryGetValue(contraext, out o))
                        d.Add(contraext, ccomtcom);
                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
               return null;
            }

        }

        public bool Agrupada(string contraext)
        {
            string o;
            return dic.TryGetValue(contraext, out o);
        }

    }
}
