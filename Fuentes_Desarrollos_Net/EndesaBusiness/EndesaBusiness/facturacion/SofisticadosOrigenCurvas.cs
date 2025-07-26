using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class SofisticadosOrigenCurvas
    {
        Dictionary<string, List<string>> dic;

        public SofisticadosOrigenCurvas()
        {
            dic = new Dictionary<string, List<string>>();
            Carga();
        }

        private void Carga()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            try
            {
                strSql = "SELECT i.cups20, toc.origen from ag_inventario i"
                    + " INNER JOIN ag_origen_curvas oc on"
                    + " oc.id_cups = i.id_cups"
                    + " INNER JOIN ag_tipos_origen_curvas toc on"
                    + " toc.id_origen = oc.id_origen"
                    + " where (i.fd <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
                    + " and i.fh >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "')"
                    + " ORDER BY i.id_cups, oc.prioridad";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    List<string> o;
                    if (!dic.TryGetValue(r["cups20"].ToString(), out o))
                    {
                        o = new List<string>();
                        o.Add(r["origen"].ToString());
                        dic.Add(r["cups20"].ToString(), o);
                    }
                    else
                        o.Add(r["origen"].ToString());
                }
                db.CloseConnection();
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public List<string> ListaPrioridades(string cups20)
        {
            List<string> o;
            if (dic.TryGetValue(cups20, out o))
                return o;
            else
                return null;
        }

    }
}
