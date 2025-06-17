using EndesaBusiness.servidores;
using EndesaEntity.medida;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace EndesaBusiness.medida.pendiente
{
    public class SegmentosPortugal_KEE : SegmentosPortugalKEE
    {
        logs.Log ficheroLog;
        Dictionary<string, List<SegmentosPortugalKEE>> dic;
        //////public bool existe { get; set; }

        public SegmentosPortugal_KEE()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "SegmentosPortugal_KEE");

            //dic = Carga(MaxFechaInforme()); //Saca el maximo de t_ed_h_pdtweb_pm_b2b
            // Hay problemas de sincronización, KEE sale todos los días, en cambio SAP no se ejecuta fines de semana y ¿Festivos?
            // A la hora de ir a buscar a KEE, buscamos la fecha maxima de SAP para sincronizar los dos
            dic = Carga();

        }

        private Dictionary<string, List<SegmentosPortugalKEE>> Carga()
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";
            MySQLDB dbAux;
            MySqlCommand commandAux;
            MySqlDataReader rAux;
            string segmento;

            MySQLDB dbgrabar;
            MySqlCommand commandgrabar;
            MySqlDataReader rgrabar;

            ////fecha_ControlCargas = MaxFechaInformeGestDiarioB2B();
            Dictionary<string, List<SegmentosPortugalKEE>> d = new Dictionary<string, List<SegmentosPortugalKEE>>();

            try
            {


                //Problema, aquí en ps_at estan solo los de vigor, si se da uno de baja desaparece
                strSql = "select A.id_crto_ext, A.cups, A.fh_alta_crto, A.fh_baja_crto, A.de_estado, A.secuencial, cd_tp_tension as SegmentoSalesforce "
                + " from ("
                + "     select id_crto_ext, cd_cups_crto as cups, fh_alta_crto, fh_baja_crto, de_estado, max(cd_sec_crto) as secuencial "
                + "     from ed_owner.t_ed_h_crtos "
                + "     where de_marca_back = 'OPERACIONES B2B' "
                + "     and(fh_baja_crto >= '2020-12-31' or fh_baja_crto is null) "
                + "     and cd_cups_ext  in "
                + "     ( "
                + "         select cd_cups_ext "
                + "         from ed_owner.t_ed_h_ps "
                + "         where lg_migrado_sap = 'S' "
                + "         and de_seg_mercado = 'SE' "
                + "         union all "
                + "         select cd_cups_ext "
                + "         from ed_owner.t_ed_h_ps "
                + "         where lg_migrado_sap = 'S' "
                + "         and cd_pais = 'PORTUGAL' "
                + "     ) "
                + "     and de_estado in ('BAJA', 'EN VIGOR') "
                + "     group by id_crto_ext ,cd_cups_crto, fh_alta_crto,fh_baja_crto,de_estado "
                + " ) as A"
                + " left join ed_owner.t_ed_h_ps"
                + " on A.cups = cd_cups_ext"
                //////+ " where  A.cups='PT0002000083670996QN'"
                + " order by A.cups asc,A.id_crto_ext asc ";

                ficheroLog.Add(strSql);
                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    SegmentosPortugalKEE c = new SegmentosPortugalKEE();
                    c.cups = r["cups"].ToString();
                    c.id_contrato = r["id_crto_ext"].ToString();
                    if (r["fh_alta_crto"] != System.DBNull.Value)
                    {
                        c.fecha_alta_cto = Convert.ToDateTime(r["fh_alta_crto"]);
                    }
                    if (r["fh_baja_crto"] != System.DBNull.Value)
                    {
                        c.fecha_baja_cto = Convert.ToDateTime(r["fh_baja_crto"]);
                    }
                    c.estado = r["de_estado"].ToString();
                    c.secuencial = r["secuencial"].ToString();

                    if (r["SegmentoSalesforce"] != System.DBNull.Value)
                    {
                        c.SegmentoSalesForce = r["SegmentoSalesforce"].ToString();
                    }

                    segmento = "";
                    if (r["cups"].ToString().Substring(0,2) == "PT")
                    {
                        // Hay que ir a buscar el segmento a la tabla de Marta (En MySql): t_ed_h_sap_pendiente_facturar_segmentos_compor (para cups de Portugal)
                        if (r["fh_baja_crto"] != System.DBNull.Value)
                        {
                            strSql = "select segmento "
                            + " from t_ed_h_sap_pendiente_facturar_segmentos_compor "
                            + " where fdesde >='" + Convert.ToDateTime(r["fh_alta_crto"]).ToString("yyyy-MM-dd") + "'"
                            + " and fdesde <='" + Convert.ToDateTime(r["fh_baja_crto"]).ToString("yyyy-MM-dd") + "'"
                            + " and cd_cups='" + r["cups"].ToString() + "'"
                            + " order by fdesde desc";
                        }
                        else
                        {
                            strSql = "select segmento "
                            + " from t_ed_h_sap_pendiente_facturar_segmentos_compor "
                            + " where fdesde <='" + Convert.ToDateTime(r["fh_alta_crto"]).ToString("yyyy-MM-dd") + "'"
                            + " and cd_cups='" + r["cups"].ToString() + "'"
                            + " order by fdesde desc";
                        }
                        dbAux = new MySQLDB(MySQLDB.Esquemas.FAC);
                        commandAux = new MySqlCommand(strSql, dbAux.con);
                        rAux = commandAux.ExecuteReader();
                       
                        while (rAux.Read())
                        {
                            if (rAux["segmento"] != System.DBNull.Value)
                                segmento = rAux["segmento"].ToString();
                            break;
                        }
                        dbAux.CloseConnection();
                    }

                    c.SegmentoCompor = segmento;
                    c.Clave = c.cups + ";" + c.fecha_alta_cto + ";" + c.estado;

                    List<SegmentosPortugalKEE> t;
                    ////// if (!d.TryGetValue(c.cups, out t))
                    //////{
                        t = new List<SegmentosPortugalKEE>();
                        t.Add(c);
                        d.Add(c.Clave, t);
                    //////}
                    //////else
                    //////    t.Add(c);
                    ///

                    ///Borrar t_ed_h_sap_pendiente_facturar_TodosSegmentos
                    strSql = "REPLACE INTO t_ed_h_sap_pendiente_facturar_TodosSegmentos (id_crto_ext, cups, fh_alta_crto, fh_baja_crto, estado, secuencial, SegmentoSalesforce, SegmentoCompor)"
                   + " values('"
                   + c.id_contrato + "','"
                   + c.cups + "',";
                   if (r["fh_alta_crto"] != System.DBNull.Value && r["fh_alta_crto"].ToString() != "01-01-0001 0:00:00")
                       strSql += "'" + Convert.ToDateTime(c.fecha_alta_cto).ToString("yyyy-MM-dd") + "',";
                   else
                       strSql += "Null,";
                   if (r["fh_baja_crto"] != System.DBNull.Value && r["fh_baja_crto"].ToString() != "01-01-0001 0:00:00")
                       strSql += "'" + Convert.ToDateTime(c.fecha_baja_cto).ToString("yyyy-MM-dd") + "','";
                   else
                       strSql += "Null,'";
                   strSql += c.estado + "','"
                   + c.secuencial + "','"
                   + c.SegmentoSalesForce + "','"
                   + c.SegmentoCompor + "')";


                    dbgrabar = new MySQLDB(MySQLDB.Esquemas.FAC);
                    commandgrabar = new MySqlCommand(strSql, dbgrabar.con);
                    commandgrabar.ExecuteNonQuery();
                    dbgrabar.CloseConnection();

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception ex)
            {
                ficheroLog.addError("Carga: " + ex.Message);
                return null;
            }
        }

        public List<string> GetCups(string cups, DateTime fecha_desde, DateTime fecha_hasta, DateTime fec_act)
        {
            List<string> lista = new List<string>();

            List<SegmentosPortugalKEE> o;
            if (dic.TryGetValue(cups, out o))
            {
                //////lista =
                //////    o.FindAll(z => z.fecha_desde == fecha_desde
                //////    && z.fecha_hasta == fecha_hasta && Convert.ToDateTime(fec_act.ToString("yyyy-MM-dd hh:mm")) > Convert.ToDateTime(z.fecha_informe.ToString("yyyy-MM-dd hh:mm"))  ).Select(z => z.estado).ToList();

                lista =
                   o.FindAll(z => z.cups == cups).Select(z => z.SegmentoCompor).ToList();

            }
            return lista;
        }

    }

}
