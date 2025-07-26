using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class Kee
    {
        public Dictionary<string, EndesaEntity.medida.Kee_Tabla> dic_starbeat;

        Dictionary<string, EndesaEntity.medida.Kee_Portugal_Tabla> dic_portugal;


        public Kee()
        {
            dic_starbeat = Carga(EndesaEntity.medida.Kee_TipoReporte.Reportes.StarBeat);
        }

        private Dictionary<string, EndesaEntity.medida.Kee_Tabla> Carga(EndesaEntity.medida.Kee_TipoReporte.Reportes reporte)
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string tabla = "";
            string[] ld;

            Dictionary<string, EndesaEntity.medida.Kee_Tabla> dic = new Dictionary<string, EndesaEntity.medida.Kee_Tabla>();

            try
            {

                switch (reporte)
                {
                    case EndesaEntity.medida.Kee_TipoReporte.Reportes.Exabeat:
                        tabla = "kee_reporte_exabeat";
                        break;
                    case EndesaEntity.medida.Kee_TipoReporte.Reportes.Ftp:
                        tabla = "kee_reporte_ftp";
                        break;
                    case EndesaEntity.medida.Kee_TipoReporte.Reportes.Publicada:
                        tabla = "kee_reporte_publicada";
                        break;
                    case EndesaEntity.medida.Kee_TipoReporte.Reportes.StarBeat:
                        tabla = "kee_reporte_starbeat";
                        break;
                }

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(QueryTablaReporte(tabla), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Kee_Tabla c = new EndesaEntity.medida.Kee_Tabla();
                    #region Campos

                    c.tiporeporte = tabla;

                    if (r["cups"] != System.DBNull.Value)
                        c.cups = r["cups"].ToString();

                    if (r["tipo_pm"] != System.DBNull.Value)
                        c.tipo_pm = r["tipo_pm"].ToString();

                    if (r["fecha_inicio"] != System.DBNull.Value)
                        c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);
                    if (r["fecha_fin"] != System.DBNull.Value)
                        c.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);
                    if (r["sumario_ae_cch"] != System.DBNull.Value)
                        c.sumatorio_ae_cch = r["sumario_ae_cch"].ToString();
                    if (r["sumatorio_r1_cch"] != System.DBNull.Value)
                        c.sumatorio_r1_cch = r["sumatorio_r1_cch"].ToString();
                    if (r["sumatorio_ae_ch"] != System.DBNull.Value)
                        c.sumatorio_ae_ch = r["sumatorio_ae_ch"].ToString();
                    if (r["sumatorio_r1_ch"] != System.DBNull.Value)
                        c.sumatorio_r1_ch = r["sumatorio_r1_ch"].ToString();

                    c.periodo_completo = r["periodo_completo"].ToString() == "S";

                    if (r["fecha_min_cch"] != System.DBNull.Value)
                        c.fecha_min_cch = Convert.ToDateTime(r["fecha_min_cch"]);
                    if (r["fecha_max_cch"] != System.DBNull.Value)
                        c.fecha_max_cch = Convert.ToDateTime(r["fecha_max_cch"]);

                    if (r["num_dias_cch"] != System.DBNull.Value)
                        c.num_dias_cch = r["num_dias_cch"].ToString();

                    if (r["fecha_min_ch"] != System.DBNull.Value)
                        c.fecha_min_cc = Convert.ToDateTime(r["fecha_min_ch"]);
                    if (r["fecha_max_ch"] != System.DBNull.Value)
                        c.fecha_max_ch = Convert.ToDateTime(r["fecha_max_ch"]);

                    if (r["num_dias_ch"] != System.DBNull.Value)
                        c.num_dias_cc = r["num_dias_ch"].ToString();

                    if (r["dias_sin_medidas_cch"] != System.DBNull.Value)
                    {
                        ld = r["dias_sin_medidas_cch"].ToString().Split(',');
                        for (int x = 0; x < ld.Count(); x++)
                            c.dias_sin_medida_cch.Add(Convert.ToDateTime(ld[x]), Convert.ToDateTime(ld[x]));
                    }

                    if (r["dias_sin_medidas_ch"] != System.DBNull.Value)
                    {
                        ld = r["dias_sin_medidas_ch"].ToString().Split(',');
                        for (int x = 0; x < ld.Count(); x++)
                            c.dias_sin_medida_ch.Add(Convert.ToDateTime(ld[x]), Convert.ToDateTime(ld[x]));
                    }

                    EndesaEntity.medida.Kee_Tabla o;
                    if (!dic.TryGetValue(c.cups, out o))
                        dic.Add(c.cups, c);


                    #endregion
                }
                db.CloseConnection();
                return dic;
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }

        }

        private string QueryTablaReporte(string tabla)
        {

            string strSql = "SELECT ml.cd_cups22 as cups, s.tdistri, s.Tipo_PM_Calculado as tipo_pm," +
                " ks.fecha_recepcion, ks.fecha_informe," +
                " ks.fecha_inicio, ks.fecha_fin, ks.sumario_ae_cch, ks.sumatorio_r1_cch, ks.sumatorio_ae_ch, ks.sumatorio_r1_ch," +
                " ks.periodo_completo, ks.fecha_min_cch, ks.fecha_max_cch, ks.num_dias_cch, ks.fecha_min_ch, ks.fecha_max_ch," +
                " ks.num_dias_ch, ks.dias_sin_medidas_cch, ks.dias_sin_medidas_ch" +
                " FROM med.dt_vw_ed_f_puntos_ml ml" +
                " LEFT OUTER JOIN med.scea s ON" +
                " s.CUPS20 = ml.cd_cups20" +
                " LEFT OUTER JOIN med." + tabla + " ks ON" +
                " ks.cups = ml.cd_cups22" +
                " WHERE ml.cd_cups22 IS NOT NULL AND ml.cd_cups22 <> ''";



            return strSql;
        }
    }
}
