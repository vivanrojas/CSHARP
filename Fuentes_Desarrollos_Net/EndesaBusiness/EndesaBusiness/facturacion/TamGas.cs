using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class TamGas
    {
        
        Dictionary<string, double> dic;
        Dictionary<int, double> dic_2;

        public TamGas()
        {
            dic_2 = Carga();
        }

        public TamGas(List<string> lista_cups20)
        {
            dic = Carga(lista_cups20);
        }

        private Dictionary<string, double> Carga(List<string> lista_cups20)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string cups20 = "";

            Dictionary<string, double> d = new Dictionary<string, double>();
            try
            {
                strSql = "select cig.CUPSREE, g.TAM from cm_inventario_gas cig inner JOIN"
                    + " tam_gas g on"
                    + " g.ID_PS = cig.ID_PS AND"
                    + " g.ID_CTO_PS = cig.ID_CTO_PS"
                    + " where cig.CUPSREE in ('" + lista_cups20[0] + "'";
                    for (int i = 1; i < lista_cups20.Count; i++)
                        strSql += ",'" + lista_cups20[i] + "'";

                    strSql += ")";
                
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    double o;
                    if (r["CUPSREE"] != System.DBNull.Value)
                    {
                        cups20 = r["CUPSREE"].ToString();
                        if (!d.TryGetValue(cups20, out o))
                            d.Add(cups20, Convert.ToDouble(r["TAM"]));
                    }

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
        private Dictionary<int, double> Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            int id_ps = 0;

            Dictionary<int, double> d = new Dictionary<int, double>();
            try
            {
                strSql = "select cig.ID_PS, g.TAM from cm_inventario_gas cig inner JOIN"
                    + " tam_gas g on"
                    + " g.ID_PS = cig.ID_PS AND"
                    + " g.ID_CTO_PS = cig.ID_CTO_PS";                    

                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    double o;                    
                    id_ps = Convert.ToInt32(r["ID_PS"]);
                    if (!d.TryGetValue(id_ps, out o))
                            d.Add(id_ps, Convert.ToDouble(r["TAM"]));
                    

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

        public double GetTamCups20(string cups20)
        {

            double o = 0;
            if (dic.TryGetValue(cups20, out o))
                return o;
            return o;

        }

        public double GetTam_ID_PS(int id_ps)
        {

            double o = 0;
            if (dic_2.TryGetValue(id_ps, out o))
                return o;
            return o;

        }
    }
    
}
