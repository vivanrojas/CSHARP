using EndesaBusiness.servidores;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.compor
{
    public class Inventario_Multipuntos
    {
        public Dictionary<string, string> dic { get; set; }

        public Inventario_Multipuntos()
        {
            dic = Carga();
        }

        private Dictionary<string, string> Carga()
        {
            OracleServer ora_db;
            OracleCommand ora_command;
            OracleDataReader r;
            string strSql = "";

            Dictionary<string, string> d = new Dictionary<string, string>();
            string cpe = "";

            try
            {
                strSql = "SELECT CPE, NOMBRE, NIF, CANAL FROM APL_INVENTARIO_MULTIPUNTOS";
                ora_db = new OracleServer(OracleServer.Servidores.COMPOR);
                ora_command = new OracleCommand(strSql, ora_db.con);
                r = ora_command.ExecuteReader();
                while (r.Read())
                {
                    if (r["CPE"] != System.DBNull.Value)
                    {
                        cpe = r["CPE"].ToString();
                        string o;
                        if (!d.TryGetValue(cpe, out o))
                            d.Add(cpe, cpe);
                    }
                    
                }
                ora_db.CloseConnection();

                return d;
            }catch(Exception e)
            {
                return null;
            }
        }

        public bool Es_MultiPunto(string cpe)
        {
            string o;
            return dic.TryGetValue(cpe, out o);
        }
    }
}
