using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion.redshift
{


    public class Clientes
    {
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;
        logs.Log ficheroLog;


        public Clientes()
        {
            ss_pp = new utilidades.Seguimiento_Procesos();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Copia_Clientes_BI");
        }

        public void CopiaClientesBI()
        {
            StringBuilder sb = new StringBuilder();
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";

            MySQLDB dbmy;
            MySqlCommand commandmy;

            bool firstOnly = true;
            int j = 0;
            int k = 0;
            int totalRegistros = 0;

            DateTime ultimaFechaCopiado = new DateTime();

            try
            {
                ss_pp.Update_Fecha_Inicio("Facturación", "Copia Clientes BI", "Copia Clientes BI");

                ultimaFechaCopiado = UltimaActualizacion();
                ficheroLog.Add(Consulta(ultimaFechaCopiado));
                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(Consulta(ultimaFechaCopiado), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    j++;
                    k++;

                    if (firstOnly)
                    {
                        sb = null;
                        sb = new StringBuilder();
                        sb.Append("REPLACE INTO t_ed_f_sap_clis");
                        sb.Append(" (cl_cli, cd_nif_cif_cli, cd_cli_ext_b2c, cd_cli_ext_b2b, cd_interloc_comer, de_interloc_comer,     ");
                        sb.Append(" de_nombre_org2, de_nombre_org3, de_nombre_org4, de_seg_nombre, cd_po_box, lg_cat_cliente, cd_clave_trat,");
                        sb.Append(" de_clave_trat, cd_forma_jur_org, lg_marca_simpl, lg_inv_no_rep_dext, nm_direccion, cd_seg_agru, lg_cli_top,");
                        sb.Append(" de_forma_jur_org, cd_func_ic, cd_usu_modif_conc, cd_med_conc, cd_usu_modif_preconc, cd_mot_descheck, de_func_ic,");
                        sb.Append(" de_usu_modif_conc, de_med_conc, de_usu_modif_preconc, de_mot_descheck, cd_pro_pago_mes, cd_pro_pago_anio,");
                        sb.Append(" cd_seg_cartera_imp, cd_seg_cartera_viva, fh_desmarc_preconc, fh_desmarc_concurso, fh_fin_preconc, fh_fin_concurso,");
                        sb.Append(" fh_marca_preconc, fh_marca_concurso, fh_preconc, fh_publica_boe, lg_asnef, fec_act, cod_carga, created_by, created_date) VALUES ");                        
                        firstOnly = false;
                    }

                    #region Campos
                    if (r["cl_cli"] != System.DBNull.Value)
                        sb.Append("('").Append(r["cl_cli"].ToString()).Append("',");
                    else
                        sb.Append("(null,");

                    if (r["cd_nif_cif_cli"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_nif_cif_cli"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_cli_ext_b2c"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_cli_ext_b2c"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_cli_ext_b2b"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_cli_ext_b2b"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_interloc_comer"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_interloc_comer"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_interloc_comer"] != System.DBNull.Value)
                        sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_interloc_comer"].ToString())).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_nombre_org2"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_nombre_org2"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_nombre_org3"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_nombre_org3"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_nombre_org4"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_nombre_org4"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_seg_nombre"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_seg_nombre"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_po_box"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_po_box"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_cat_cliente"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_cat_cliente"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_clave_trat"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_clave_trat"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_clave_trat"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_clave_trat"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_forma_jur_org"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_forma_jur_org"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_marca_simpl"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_marca_simpl"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_inv_no_rep_dext"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_inv_no_rep_dext"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_direccion"] != System.DBNull.Value)
                        sb.Append("'").Append(r["nm_direccion"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_seg_agru"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_seg_agru"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_cli_top"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_cli_top"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_forma_jur_org"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_forma_jur_org"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_func_ic"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_func_ic"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_usu_modif_conc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_usu_modif_conc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_med_conc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_med_conc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_usu_modif_preconc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_usu_modif_preconc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_mot_descheck"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_mot_descheck"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_func_ic"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_func_ic"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_usu_modif_conc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_usu_modif_conc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_med_conc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_med_conc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_usu_modif_preconc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_usu_modif_preconc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_mot_descheck"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_mot_descheck"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_pro_pago_mes"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_pro_pago_mes"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_pro_pago_año"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_pro_pago_año"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_seg_cartera_imp"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_seg_cartera_imp"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_seg_cartera_viva"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_seg_cartera_viva"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_desmarc_preconc"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_desmarc_preconc"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_desmarc_concurso"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_desmarc_concurso"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_fin_preconc"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_fin_preconc"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_fin_concurso"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_fin_concurso"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_marca_preconc"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_marca_preconc"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_marca_concurso"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_marca_concurso"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_preconc"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_preconc"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_publica_boe"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_publica_boe"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");                    

                    if (r["lg_asnef"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_asnef"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fec_act"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fec_act"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cod_carga"] != System.DBNull.Value)
                        sb.Append(r["cod_carga"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                    #endregion

                    if (j == 50)
                    {
                        Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                        firstOnly = true;
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                        commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        j = 0;
                    }


                }
                db.CloseConnection();

                if (j > 0)
                {
                    Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                    firstOnly = true;
                    dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                    commandmy.ExecuteNonQuery();
                    dbmy.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    j = 0;
                }

                ss_pp.Update_Fecha_Fin("Facturación", "Copia Clientes BI", "Copia Clientes BI");
            }
            catch(Exception ex)
            {
                ficheroLog.addError(ex.Message);
            }
        }

        public DateTime UltimaActualizacion()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime fecha = new DateTime();

            fecha = new DateTime(2022, 01, 01);

            strSql = "SELECT max(fec_act) AS fec_act FROM t_ed_f_sap_clis";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["fec_act"] != System.DBNull.Value)
                    fecha = Convert.ToDateTime(r["fec_act"]);
            }

            db.CloseConnection();
            Console.WriteLine("Última fecha de copiado RedShift: " + fecha.ToString("dd/MM/yyyy"));
            ficheroLog.Add("Última fecha de copiado RedShift: " + fecha.ToString("dd/MM/yyyy"));

            return fecha;
        }
        public string Consulta(DateTime fecha)
        {
            string strSql = "";

            strSql = "SELECT cl_cli, cd_nif_cif_cli, cd_cli_ext_b2c, cd_cli_ext_b2b, cd_interloc_comer,"
                + " de_interloc_comer, de_nombre_org2, de_nombre_org3, de_nombre_org4, de_seg_nombre,"
                + " cd_po_box, lg_cat_cliente, cd_clave_trat, de_clave_trat, cd_forma_jur_org, lg_marca_simpl,"
                + " lg_inv_no_rep_dext, nm_direccion, cd_seg_agru, lg_cli_top, de_forma_jur_org, cd_func_ic,"
                + " cd_usu_modif_conc, cd_med_conc, cd_usu_modif_preconc, cd_mot_descheck, de_func_ic,"
                + " de_usu_modif_conc, de_med_conc, de_usu_modif_preconc, de_mot_descheck, cd_pro_pago_mes,"
                + " cd_pro_pago_año, cd_seg_cartera_imp, cd_seg_cartera_viva, fh_desmarc_preconc, fh_desmarc_concurso,"
                + " fh_fin_preconc, fh_fin_concurso, fh_marca_preconc, fh_marca_concurso, fh_preconc, fh_publica_boe,"
                + " lg_asnef, fec_act, cod_carga"
                + " FROM ed_owner.t_ed_f_sap_clis where"
                + " fec_act >= '" + fecha.AddDays(-1).ToString("yyyy-MM-dd") + "' AND"
                + " cd_cli_ext_b2b is not null";

            return strSql;
        }

    }
}
