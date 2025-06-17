using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class Anomalias : EndesaEntity.facturacion.Anomalia
    {
        Dictionary<string, EndesaEntity.facturacion.Anomalia> dic;
        public Anomalias()
        {
            dic = Carga();
        }

        public Dictionary<string, EndesaEntity.facturacion.Anomalia> Carga()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.facturacion.Anomalia> d =
                new Dictionary<string, EndesaEntity.facturacion.Anomalia>();

            DateTime maxFecha = new DateTime();
            try
            {
                strSql = "SELECT MAX(f.fejecucion) maxFecha FROM fact.fo_agora f";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                    maxFecha = Convert.ToDateTime(r["maxFecha"]);
                db.CloseConnection();

                strSql = "SELECT /*+ MAX_EXECUTION_TIME(500000)*/"
                    + " f.cups22,"
                    + " f.anomal1, f.anomal2, f.anomal3, f.anomal4, f.anomal5"
                    + " FROM fact.fo_agora f" 
                    + " WHERE f.fejecucion = '" + maxFecha.ToString("yyyy-MM-dd") + "'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.Anomalia c = new EndesaEntity.facturacion.Anomalia();
                    c.cups20 = r["cups22"].ToString().Substring(0, 20);
                    c.existe = true;
                    for (int i = 0; i < 5; i++)
                        if (r["anomal" + (i + 1)] != System.DBNull.Value)
                            c.anomalia[i] = r["anomal" + (i + 1)].ToString();

                    EndesaEntity.facturacion.Anomalia o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);

                }                    
                db.CloseConnection();


                return d;
            }
            catch(Exception ex)
            {
                return null;
            }
                       
        }

        public void GetAnomalia(string cups20)
        {
            EndesaEntity.facturacion.Anomalia o;
            if (dic.TryGetValue(cups20, out o))
            {
                this.existe = true;
                this.cups20 = o.cups20;
                this.anomalia = o.anomalia;
            }
            else
            {
                this.existe = false;
            }
        }

    }
}
