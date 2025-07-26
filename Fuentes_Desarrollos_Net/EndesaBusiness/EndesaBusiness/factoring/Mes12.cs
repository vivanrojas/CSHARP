using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.factoring
{
    public class Mes12
    {
        public Dictionary<string, string> dic_facturas { get; set; }
        public Mes12()
        {
            dic_facturas = CargaDatos();
        }

        private Dictionary<string, string> CargaDatos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string cfactura = "";

            Dictionary<string, string> d = new Dictionary<string, string>();
            try
            {

                strSql = "SELECT m.HOJA, m.CFACTURA FROM mes12_excel m"
                    + " WHERE m.HOJA NOT LIKE 'NO%' AND"
                    + " m.TIPO IS NULL";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["CFACTURA"] != System.DBNull.Value)
                        cfactura = r["CFACTURA"].ToString();

                    string o;
                    if (!d.TryGetValue(cfactura, out o))
                        d.Add(cfactura, cfactura);
                }
                db.CloseConnection();
                return d;

            }
            catch (Exception ex)
            {
                return null;
            }
        }


        public bool ExisteFacturaEnMes12(string cfactura)
        {
            string o;
            return dic_facturas.TryGetValue(cfactura, out o);
        }
    }
}
