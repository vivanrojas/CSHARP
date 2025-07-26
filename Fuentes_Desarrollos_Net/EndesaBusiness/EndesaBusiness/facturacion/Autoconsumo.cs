using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class Autoconsumo
    {
        public Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Autoconsumo> dic;
        public Autoconsumo()
        {
            dic = Carga();
        }

        private Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Autoconsumo> Carga()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Autoconsumo> d =
                new Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Autoconsumo>();

            try
            {
                strSql = "SELECT /*+ MAX_EXECUTION_TIME(500000)*/ ps.EMPRESA, ps.NIF, ps.Cliente,"
                    + " a.cups22 , ps.tarifa, p.descripcion, ps.MIGRADO_SAP"
                    + " FROM fact.cm_autoconsumo_param p"
                    + " INNER JOIN cont.ps_autoconsumos a ON"
                    + " p.id = a.complemento_autoconsumo"
                    + " INNER JOIN cont.PS_AT ps ON"
                    + " substr(ps.cups22, 1, 20) = substr(a.cups22, 1, 20)"
                    + " GROUP BY a.cups22;";

                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.cuadroDeMando.Autoconsumo c = new EndesaEntity.facturacion.cuadroDeMando.Autoconsumo();
                    c.origen_sistemas = r["MIGRADO_SAP"].ToString() == "S" ? "SAP" : "SCE";
                    c.empresa = r["EMPRESA"].ToString();
                    c.nif = r["NIF"].ToString();
                    c.cliente = r["Cliente"].ToString();
                    c.cups22 = r["cups22"].ToString();
                    c.tarifa = r["tarifa"].ToString();
                    c.descripcion = r["descripcion"].ToString();

                    EndesaEntity.facturacion.cuadroDeMando.Autoconsumo o;
                    if (!d.TryGetValue(c.cups22, out o))
                        d.Add(c.cups22, c);

                }
                db.CloseConnection();

                return d;
            }catch(Exception ex)
            {
                return null;
            }
        }


    }
}
