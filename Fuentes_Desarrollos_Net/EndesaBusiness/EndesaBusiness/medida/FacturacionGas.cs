using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class FacturacionGas
    {
        Dictionary<int, EndesaEntity.facturacion.FacturaGas> dic;
        public FacturacionGas()
        {
            //ActualizaTablaUltimaFacturaEmitida();
            dic = CargaFacturas();
        }


        private Dictionary<int, EndesaEntity.facturacion.FacturaGas> CargaFacturas()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            Dictionary<int, EndesaEntity.facturacion.FacturaGas> d = new Dictionary<int, EndesaEntity.facturacion.FacturaGas>();

            try
            {
                Console.WriteLine("Cargando datos de fact.fo_s_ultima_factura_emitida");
                strSql = "select ID_PS, FH_FACTURA, FH_INI_FACTURACION, FH_FIN_FACTURACION, CD_NFACTURA_REALES_PS"
                    + " from fo_s_ultima_factura_emitida";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.FacturaGas c = new EndesaEntity.facturacion.FacturaGas();

                    c.id_ps = Convert.ToInt32(r["ID_PS"]);
                    if (r["FH_FACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FH_FACTURA"]);
                    if (r["FH_INI_FACTURACION"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["FH_INI_FACTURACION"]);
                    if (r["FH_FIN_FACTURACION"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["FH_FIN_FACTURACION"]);


                    EndesaEntity.facturacion.FacturaGas o;
                    if (!d.TryGetValue(c.id_ps, out o))
                        d.Add(c.id_ps, c);
                }
                db.CloseConnection();

                return d;

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }


        private void ActualizaTablaUltimaFacturaEmitida()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            strSql = "REPLACE INTO fo_s_ultima_factura_emitida"
                + " SELECT s.*"
                + " FROM fo_s s"
                + " WHERE s.FH_ULT_ACTUALIZACION IS NOT NULL AND"
                + " s.FH_FACTURA >= s.FH_FIN_FACTURACION AND"
                + " (instr(s.CD_NFACTURA_REALES_PS, 'S') < 1 AND instr(s.CD_NFACTURA_REALES_PS, 'A') < 1)"
                + " ORDER BY s.ID_PS, s.FH_ULT_ACTUALIZACION";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        public int UltimoPeriodoFacturado(int id_ps)
        {
            EndesaEntity.facturacion.FacturaGas o;
            if (dic.TryGetValue(id_ps, out o))
            {
                return Convert.ToInt32(o.ffactdes.ToString("yyyyMM"));
            }
            else
                return 0;
        }


    }
}
