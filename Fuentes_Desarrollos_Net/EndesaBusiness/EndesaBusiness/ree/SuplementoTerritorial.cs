using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.ree
{
    class SuplementoTerritorial
    {

        Dictionary<string, List<string>> dic;
        EndesaBusiness.facturacion.SuplementosTerritoriales sup;
        public SuplementoTerritorial()
        {
            sup = new facturacion.SuplementosTerritoriales();
            dic = Carga();
        }


        public Dictionary<string, List<string>> Carga()
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            string cups20 = "";
            string texto = "";

            Dictionary<string, List<string>> d = new Dictionary<string, List<string>>();

            try
            {
                strSql = "SELECT t.cups20, t.texto FROM eer_suplemento_territorial t"
                    + " ORDER BY t.cups20, t.linea";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    cups20 = r["cups20"].ToString();
                    texto = r["texto"].ToString();

                    List<string> o;
                    if (!d.TryGetValue(cups20, out o))
                    {
                        o = new List<string>();
                        o.Add(texto);
                        d.Add(cups20, o);
                    }
                    else
                        o.Add(texto);

                }
                db.CloseConnection();
                return d;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        public List<string> GetComplementoTerritorial(string cups20)
        {
            List<string> o;
            if (dic.TryGetValue(cups20, out o))
            {
                return o;
            }                
            else
                return null;

        }

        public int CuotasPdtes(string cups20, int yyyymm)
        {
            return sup.NumCuotas(cups20, yyyymm);
        }

    }
}
