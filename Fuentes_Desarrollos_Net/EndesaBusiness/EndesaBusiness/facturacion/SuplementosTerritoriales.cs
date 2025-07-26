using EndesaBusiness.calendarios;
using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class SuplementosTerritoriales
    {
        Dictionary<string, List<EndesaEntity.facturacion.SuplementoTerritorial>> dic;
        public SuplementosTerritoriales()
        {
            dic = Carga();
        }


        private Dictionary<string, List<EndesaEntity.facturacion.SuplementoTerritorial>> Carga()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, List<EndesaEntity.facturacion.SuplementoTerritorial>> d 
                = new Dictionary<string, List<EndesaEntity.facturacion.SuplementoTerritorial>>();

            try
            {

                strSql = "select cups20, mes, cuotas_restantes from eer_cuotas_sstt";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.SuplementoTerritorial c = new EndesaEntity.facturacion.SuplementoTerritorial();
                    if (r["cups20"] != System.DBNull.Value)
                        c.cups20 = r["cups20"].ToString();
                    if (r["mes"] != System.DBNull.Value)
                        c.aniomes = Convert.ToInt32(r["mes"]);
                    if (r["cuotas_restantes"] != System.DBNull.Value)
                        c.num_cuota = Convert.ToInt32(r["cuotas_restantes"]);

                    List<EndesaEntity.facturacion.SuplementoTerritorial> o;
                    if (!d.TryGetValue(c.cups20, out o))
                    {
                        o = new List<EndesaEntity.facturacion.SuplementoTerritorial>();
                        o.Add(c);
                        d.Add(c.cups20, o);
                    }
                    else
                        o.Add(c);

                }
                db.CloseConnection();

                return d;
            }
            catch(Exception e)
            {
                return null;
            }
            
        }

        public int NumCuotas(string cups20, int aniomes)
        {
            int numCuotas = 0;

            List<EndesaEntity.facturacion.SuplementoTerritorial> o;
            if (dic.TryGetValue(cups20, out o))
                foreach(EndesaEntity.facturacion.SuplementoTerritorial p in o)
                {
                    if (p.aniomes == aniomes)
                        return p.num_cuota;
                }

            return numCuotas;


        }


    }
}
