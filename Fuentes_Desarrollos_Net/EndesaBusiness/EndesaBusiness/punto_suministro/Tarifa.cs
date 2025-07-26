using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.punto_suministro
{
    public class Tarifa
    {
        public Dictionary<string, EndesaEntity.punto_suministro.Tarifa> dic { get; set; }

        public Tarifa()
        {
            dic = Carga();
        }

        public Tarifa(DateTime fd, DateTime fh)
        {
            dic = Carga(fd, fh);
        }

        private Dictionary<string, EndesaEntity.punto_suministro.Tarifa> Carga()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.punto_suministro.Tarifa> d
                = new Dictionary<string, EndesaEntity.punto_suministro.Tarifa>();

            try
            {
                strSql = "SELECT descripcion as Tarifa, Descripcion, "
                    + " periods_energy as NumPeriodosTarifarios"
                    + " FROM fact.fact_param_codigos_tarifas_atr";                
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.punto_suministro.Tarifa c = new EndesaEntity.punto_suministro.Tarifa();
                    c.tarifa = r["Tarifa"].ToString();
                    c.numPeriodosTarifarios = Convert.ToInt32(r["NumPeriodosTarifarios"]);
                    d.Add(c.tarifa, c);
                }
                db.CloseConnection();
                return d;
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }


        }


        private Dictionary<string, EndesaEntity.punto_suministro.Tarifa> Carga(DateTime fd, DateTime fh)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.punto_suministro.Tarifa> d
                = new Dictionary<string, EndesaEntity.punto_suministro.Tarifa>();

            try
            {
                strSql = "SELECT descripcion as Tarifa, Descripcion, "
                    + " periods_energy as NumPeriodosTarifarios"
                    + " FROM fact.fact_param_codigos_tarifas_atr where"
                    + " from_date <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " to_date >= '" + fh.ToString("yyyy-MM-dd") + "'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.punto_suministro.Tarifa c = new EndesaEntity.punto_suministro.Tarifa();
                    c.tarifa = r["Tarifa"].ToString();
                    c.numPeriodosTarifarios = Convert.ToInt32(r["NumPeriodosTarifarios"]);
                    d.Add(c.tarifa, c);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }


        }


        public EndesaEntity.punto_suministro.Tarifa GetTarifa(string tarifa)
        {
            EndesaEntity.punto_suministro.Tarifa o;
            if (dic.TryGetValue(tarifa, out o))
                return o;
            return null;
           
        }
    }
}
