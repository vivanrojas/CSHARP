using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class HistoricoCuadroMandoFacturacion : EndesaEntity.facturacion.DetalleCuadroMando
    {
        Dictionary<string, EndesaEntity.facturacion.DetalleCuadroMando> dic;
        public HistoricoCuadroMandoFacturacion(List<string> lista_cups20)        {
            
            dic = Carga(lista_cups20);
        }

        private Dictionary<string, EndesaEntity.facturacion.DetalleCuadroMando> Carga(List<string> lista_cups20)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            DateTime maxFecha = new DateTime();

            Dictionary<string, EndesaEntity.facturacion.DetalleCuadroMando> d =
                new Dictionary<string, EndesaEntity.facturacion.DetalleCuadroMando>();

            try
            {

                strSql = "select max(F_ULT_MOD) as maxFecha from cm_detalle_hist";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                    maxFecha = Convert.ToDateTime(r["maxFecha"]);
                db.CloseConnection();

                //strSql = "select a.EmpresaTitular, a.PM, a.CodContrato, b.aaaammPdte, a.Distribuidora, a.Estado, a.Subestado"
                //+ " from cm_pendiente_ml a"
                //+ " inner join (select PM, max(aaaammPdte) aaaammPdte from cm_pendiente_ml where "
                //+ " substr(PM,1,13) in ("
                //+ "'" + lista_cups13[0] + "'";

                //for (int i = 1; i < lista_cups13.Count; i++)
                //    strSql += ",'" + lista_cups13[i] + "'";

                //strSql += ") group by substr(PM,1,13)) b on"
                //+ " b.PM = a.PM and b.aaaammPdte = a.aaaammPdte where"
                //+ " substr(a.PM,1,13) in ("
                //+ "'" + lista_cups13[0] + "'";

                //for (int i = 1; i < lista_cups13.Count; i++)
                //    strSql += ",'" + lista_cups13[i] + "'";

                //strSql += ") group by substr(a.PM,1,13) order by a.aaaammPdte";

                strSql = "SELECT CNIFDNIC, DAPERSOC, IdPS, CCOUNIPS, CUPSREE, EstadoContrato, Provincia," 
                    + " UltimoMesFacturado, Mes, ID_ESTADO, Estado, Comentario, PromedioFacturacion, Grupo,"
                    + " FINICIO, FFIN, Tipo, F_ULT_MOD"
                    + " FROM fact.cm_detalle_hist where "
                    + " CUPSREE in (" + "'" + lista_cups20[0] + "'";

                for (int i = 1; i < lista_cups20.Count; i++)
                    strSql += ",'" + lista_cups20[i] + "'";

                strSql += ") and"
                    + " F_ULT_MOD = '" + maxFecha.ToString("yyyy-MM-dd") + "'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.DetalleCuadroMando c = new EndesaEntity.facturacion.DetalleCuadroMando();
                    c.nif = r["CNIFDNIC"].ToString();
                    if (r["DAPERSOC"] != System.DBNull.Value)
                        c.nombre_cliente = r["DAPERSOC"].ToString();
                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cups20 = r["CUPSREE"].ToString();

                    if(r["UltimoMesFacturado"] != System.DBNull.Value)
                        c.ultimo_mes_facturado = Convert.ToInt32(r["UltimoMesFacturado"]);

                    if (r["Estado"] != System.DBNull.Value)
                        c.estado_ltp = r["Estado"].ToString();

                    if (r["Tipo"] != System.DBNull.Value)
                        c.tipo = r["Tipo"].ToString();

                    EndesaEntity.facturacion.DetalleCuadroMando o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);
                }
                db.CloseConnection();
                return d;

            }
            catch (Exception e)
            {
                //ficheroLog.AddError("CargaPendiente: " + e.Message);
                return null;
            }
        }

        public bool Existe_CUSP20(string cups20)
        {
            bool existe = false;
            EndesaEntity.facturacion.DetalleCuadroMando o;
            if(dic.TryGetValue(cups20, out o))
            {
                existe = true;
                this.nif = o.nif;
                this.tipo = o.tipo;
                this.estado_ltp = o.estado_ltp;
                this.ultimo_mes_facturado = o.ultimo_mes_facturado;
                
            }
            return existe;
        }

    }
}
