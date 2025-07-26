using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using EndesaBusiness.calendarios;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion.cuadro_mando
{
    public class Riesgo
    {
        Dictionary<string, string> dic;

        public Riesgo()
        {
            dic = Carga();
        }

        private Dictionary<string, string> Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            StringBuilder sb = new StringBuilder();


            Dictionary<string, string> d
                = new Dictionary<string, string>();

            string nif;

            try
            {
                sb.Append("SELECT CIF FROM ccmm_riesgo ");
                sb.Append(" WHERE (FECHA_DESDE <= '").Append(DateTime.Now.ToString("yyyy-MM-dd")).Append("' AND");
                sb.Append(" FECHA_HASTA >= '").Append(DateTime.Now.ToString("yyyy-MM-dd")).Append("')");
                sb.Append(" GROUP BY CIF");
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    
                    nif = r["CIF"].ToString();                      
                    d.Add(nif, nif);
                }

                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public bool Es_Riesgo(string nif)
        {
            string o;
            return (dic.TryGetValue(nif, out o));
                
        }
    }
}
