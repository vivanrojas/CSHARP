using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using EndesaBusiness.calendarios;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion.cuadro_mando
{
    public class GrupoCliente: EndesaEntity.facturacion.cuadroDeMando.GrupoCliente
    {

        logs.Log ficheroLog;
        public Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.GrupoCliente> dic { get; set; }

        public GrupoCliente()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Informe_CuadroDeMando");
            dic = Carga();
        }

        private Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.GrupoCliente> Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            

            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.GrupoCliente> d =
                new Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.GrupoCliente>();
            try
            {
                strSql = "SELECT grupo, cliente, cups20, fecha_desde, fecha_hasta,"
                    + " created_by, created_date, last_update_by, last_update_date"
                    + " FROM fact.ccmm_grupo_clientes where"
                    + " fecha_desde <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_hasta >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.cuadroDeMando.GrupoCliente c =
                        new EndesaEntity.facturacion.cuadroDeMando.GrupoCliente();

                    c.grupo = r["grupo"].ToString();
                    c.cliente = r["cliente"].ToString();
                    c.cups20 = r["cups20"].ToString();
                    c.fecha_desde = Convert.ToDateTime(r["fecha_desde"]);
                    c.fecha_hasta = Convert.ToDateTime(r["fecha_hasta"]);

                    EndesaEntity.facturacion.cuadroDeMando.GrupoCliente o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("GrupoCliente.Carga: " + ex.Message);
                return null;
            }
        }

        
    }
}
