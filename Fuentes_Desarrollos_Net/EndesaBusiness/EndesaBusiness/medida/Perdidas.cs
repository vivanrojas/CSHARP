using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    class Perdidas
    {
        public Dictionary<string, EndesaEntity.medida.Perdidas> dic;
        public Perdidas()
        {
            dic = Carga();
        }

        private Dictionary<string, EndesaEntity.medida.Perdidas> Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            Dictionary<string, EndesaEntity.medida.Perdidas> d =
                new Dictionary<string, EndesaEntity.medida.Perdidas>();

            try
            {
                strSql = "SELECT c.CUPS13, p.PERDPACT, p.VPOTTRAN" 
                    + " FROM p_cups c"
                    + " INNER JOIN p_perdidas_ml p ON"
                    + " p.CUPS_ID = c.CUPS_ID"
                    + " WHERE p.TMEDBAJA = 'S'"
                    + " GROUP BY c.CUPS13";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Perdidas c = new EndesaEntity.medida.Perdidas();   
                    if(r["CUPS13"] != System.DBNull.Value)
                    {
                        c.cups13 = r["CUPS13"].ToString();
                        if (r["PERDPACT"] != System.DBNull.Value)
                            c.perdidas_pactadas = Convert.ToDouble(r["PERDPACT"]);
                        if (r["VPOTTRAN"] != System.DBNull.Value)
                            c.valor_potencia_transformador = Convert.ToDouble(r["VPOTTRAN"]);

                        d.Add(c.cups13, c);
                    }
                    

                }
                db.CloseConnection();
                return d;

            }
            catch(Exception e)
            {
                return null;
            }
        }

        public bool Medido_En_Baja(string cups13)
        {
            EndesaEntity.medida.Perdidas o;
            return dic.TryGetValue(cups13, out o);            
        }
    }
}
