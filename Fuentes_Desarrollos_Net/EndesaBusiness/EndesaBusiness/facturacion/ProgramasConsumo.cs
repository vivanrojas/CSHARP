using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class ProgramasConsumo
    {
        Dictionary<string,List<EndesaEntity.facturacion.ProgramasConsumo>> dic;
        public ProgramasConsumo(DateTime fd, DateTime fh)
        {
            dic = Carga(null, fd, fh);
        }

        public ProgramasConsumo(string cliente, DateTime fd, DateTime fh)
        {
            dic = Carga(cliente, fd, fh);
        }

        private Dictionary<string, List<EndesaEntity.facturacion.ProgramasConsumo>> Carga(string vcliente, DateTime fd, DateTime fh)
        {

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string cliente = "";

            Dictionary<string, List<EndesaEntity.facturacion.ProgramasConsumo>> d 
                = new Dictionary<string, List<EndesaEntity.facturacion.ProgramasConsumo>>();


            try
            {
                strSql = "SELECT pc.CUPSREE, pc.Cliente, p.*"
                    + " FROM fact.ag_programas p"
                    + " LEFT OUTER JOIN fact.ag_programascliente pc ON"
                    + " pc.CCOUNIPS = p.CCOUNIPS"
                    + " where";

                if (vcliente != null)
                    strSql = " Cliente = '" + vcliente + "' and";

                strSql +=  " (Fecha >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " Fecha <= '" + fh.ToString("yyyy-MM-dd") + "')";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.ProgramasConsumo c = new EndesaEntity.facturacion.ProgramasConsumo();
                    cliente = r["Cliente"].ToString();
                    c.cups20 = r["CUPSREE"].ToString().Substring(0,20);
                    c.fecha = Convert.ToDateTime(r["Fecha"]);
                    c.mercado = Convert.ToInt32(r["Mercado"]);
                    c.unidad = r["Unidad"].ToString();

                    for (int i = 1; i <= 25; i++)
                        if (r["Value" + i] != System.DBNull.Value)
                            c.value[i] = Convert.ToDouble(r["Value" + i]);

                    List<EndesaEntity.facturacion.ProgramasConsumo> o;
                    if(!d.TryGetValue(cliente, out o))
                    {
                        o = new List<EndesaEntity.facturacion.ProgramasConsumo>();
                        o.Add(c);
                        d.Add(cliente, o);
                    }
                    else
                    {
                        o.Add(c);
                    }

                }
                db.CloseConnection();
                return d;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return d;
            }
        }


    }
}
