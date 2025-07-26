using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot.Types;

namespace EndesaBusiness.contratacion.Redshift
{
    public class Solicitudes
    {
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;
        logs.Log ficheroLog;
        EndesaBusiness.utilidades.Param param;
        


        public Solicitudes()
        {
            param = new EndesaBusiness.utilidades.Param("t_ed_h_sol_atr_param", MySQLDB.Esquemas.CON);
            ss_pp = new utilidades.Seguimiento_Procesos();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_Copia_SOL_ATR_ROSETTA");
            
        }


        public void CopiaDatos()
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
            bool hay_error = true;

            DateTime ultima_actualizacion = new DateTime();

            try
            {

                ultima_actualizacion = Convert.ToDateTime(param.GetValue("ultima_actualizacion"));

                if (param.GetValue("estado_proceso") != "en ejecución" && (ultima_actualizacion.Date < DateTime.Now.Date || param.GetValue("estado_proceso") == "lanzar"))
                {
                    param.UpdateParameter("estado_proceso", "en ejecución");

                    ss_pp.Update_Fecha_Inicio("Contratación", "Copia Inventario BI PSAT_ROSETTA",
                        "Copia Inventario BI PSAT_ROSETTA");

                    //strSql = "delete from t_ed_h_ps_tmp";
                    //ficheroLog.Add(strSql);
                    //dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    //commandmy = new MySqlCommand(strSql, dbmy.con);
                    //commandmy.ExecuteNonQuery();
                    //dbmy.CloseConnection();

                    //ultima_actualizacion = UltimaActualizacion();

                    ficheroLog.Add(Consulta());
                    db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                    command = new OdbcCommand(Consulta(), db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        j++;
                        k++;

                        if (firstOnly)
                        {
                            sb.Append("REPLACE INTO t_ed_h_sol_atr");
                            sb.Append("(id, cd_empr, de_empr, cd_cups, id_solicitud_atr, req_name, cd_estado_sol,");
                            sb.Append("de_estado_sol, cd_tipo_sol, de_tipo_sol, cd_sub_tipo_sol, de_sub_tipo_sol,");
                            sb.Append("cd_tipo_motivo_baja, de_tipo_motivo_baja, cd_motivo_denuncia,");
                            sb.Append("de_motivo_denuncia, fh_creacion, fh_envio, fh_real_envio, fh_aceptacion, fh_rechazo,");
                            sb.Append("fh_serv_crto, cd_nif_cif_req, cd_tp_cli, cd_nif_cif_cli, tx_nom_cli, tx_apell_cli,");
                            sb.Append("de_seg_mercado, cd_linea_negocio, de_dir_cups, cd_empr_distdora, de_empr_distdora,");
                            sb.Append("name_empr_distdora, cd_crto_atr, de_comentario_sol1, cd_mot_anulacion, de_mot_anulacion,");
                            sb.Append("cd_mot_baja, de_mot_baja, nm_pot_ctatada_1, nm_pot_ctatada_2, nm_pot_ctatada_3,");
                            sb.Append("nm_pot_ctatada_4, nm_pot_ctatada_5, nm_pot_ctatada_6, cd_tarifa_contratar,");
                            sb.Append("de_tarifa_contratar, cd_tension_contratar, de_tension_contratar, cd_tp_tension,");
                            sb.Append("de_tp_tension, cd_crto_ps, cd_ver_crto_ps, cd_tp_crto_ps, cd_estado_crto_ps, cd_uso_ps,");
                            sb.Append("de_nom_gestor, lg_manual, lg_medida_baja, nm_pot_transform, nm_porcentaje_perd,");
                            sb.Append("md_fh_efecto_solic, md_fh_efecto_rec, fh_efecto_solic, fh_efecto_rec, lg_corte_crto,");
                            sb.Append("lg_reenganche_crto, nm_total_rechazos,");

                            for(int i = 1; i <= 10; i++)
                            {
                                sb.Append("cd_mot_rechazo_" + i).Append(",");
                                sb.Append("de_mot_rechazo_" + i).Append(",");
                                sb.Append("de_comentario_mot_rechazo" + i).Append(",");
                            }

                            sb.Append("tx_datos_personales, md_fh_solicitud, md_fh_respuesta, fh_respuesta, cd_tarifa,"); 
                            sb.Append("de_tarifa, cd_tp_modificacion, de_tp_modificacion, lg_modifyxml, lg_noenvdistribuidora,");
                            sb.Append("cd_tipo_sol_adm, de_tipo_sol_adm, cd_cnae, de_cnae, fh_env_sistema, fh_envio_anulacion,");
                            sb.Append("fh_aceptacion_anulacion, fh_rechazo_anulacion, cd_tipo_crto, de_tipo_crto,");
                            sb.Append("cd_tipo_crto_contratar,de_tipo_crto_contratar, atr_gestionado_cliente,");
                            sb.Append("fh_real_recep_aceptacion, cd_potencia_cont, cd_potencia_cont_2, cd_ciclo, cd_periodos,");
                            sb.Append("nm_fases, de_tipo_instalacion, lg_cliente_prioritario, cd_cae_actual, cd_cae_cliente,");
                            sb.Append("de_territory, de_canal_entrada, cd_de_reclamacion_sve, cd_estado_reclamacion,");
                            sb.Append("fh_real_recep_rechazo, de_tipo_rechazo, lg_rechazo_tras_aceptacion, lg_rechazo_manual,");
                            sb.Append("fh_real_recep_activacion, fh_real_recep_acep_anul, fh_real_recep_rech_anul, cd_tension,");
                            sb.Append("de_tension, fh_finalizacion_contratar, fh_finalizacion, fh_act_dmco, fh_real_env_anul, fh_anul,");
                            sb.Append("cd_motivo_cb_atr, de_motivo_cb_atr, id_documento) values ");

                            firstOnly = false;

                        }
                        #region Campos
                        if (r["id"] != System.DBNull.Value)
                            sb.Append("('").Append(r["id"].ToString()).Append("',");
                        else
                            sb.Append("(null,");

                        if (r["cd_pais"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_pais"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_empr"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_empr"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_empr"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_empr"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_cups"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_cups"].ToString()).Append("',");
                        else
                            sb.Append("null,");                       

                        if (r["id_solicitud_atr"] != System.DBNull.Value)
                            sb.Append("'").Append(r["id_solicitud_atr"].ToString()).Append("',");
                        else
                            sb.Append("null,");                       

                        if (r["req_name"] != System.DBNull.Value)
                            sb.Append("'").Append(r["req_name"].ToString()).Append("',");
                        else
                            sb.Append("null,");                        

                        if (r["cd_estado_sol"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_estado_sol"].ToString()).Append("',");
                        else
                            sb.Append("null,");
                        
                        if (r["de_estado_sol"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_estado_sol"].ToString()).Append("',");
                        else
                            sb.Append("null,");                        

                        if (r["cd_tipo_sol"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tipo_sol"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_tipo_sol"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_tipo_sol"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_sub_tipo_sol"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_sub_tipo_sol"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_sub_tipo_sol"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_sub_tipo_sol"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tipo_motivo_baja"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tipo_motivo_baja"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_tipo_motivo_baja"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_tipo_motivo_baja"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_motivo_denuncia"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_motivo_denuncia"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_motivo_denuncia"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_motivo_denuncia"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_creacion"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_creacion"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_envio"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_envio"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_real_envio"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_real_envio"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_aceptacion"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_aceptacion"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_rechazo"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_rechazo"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_serv_crto"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_serv_crto"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_nif_cif_req"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_nif_cif_req"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_cli"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_cli"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_nif_cif_cli"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_nif_cif_cli"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["tx_nom_cli"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tx_nom_cli"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["tx_apell_cli"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tx_apell_cli"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_seg_mercado"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_seg_mercado"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_linea_negocio"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_linea_negocio"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_dir_cups"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_dir_cups"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_empr_distdora"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_empr_distdora"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_empr_distdora"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_empr_distdora"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["name_empr_distdora"] != System.DBNull.Value)
                            sb.Append("'").Append(r["name_empr_distdora"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_crto_atr"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_crto_atr"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_comentario_sol1"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_comentario_sol1"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_mot_anulacion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_mot_anulacion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_mot_anulacion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_mot_anulacion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_mot_baja"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_mot_baja"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_mot_baja"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_mot_baja"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        for (int i = 1; i <= 6; i++)
                        {
                            if (r["nm_pot_ctatada_" + i] != System.DBNull.Value)
                                sb.Append(r["nm_pot_ctatada_" + i].ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");
                        }

                        if (r["cd_tarifa_contratar"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tarifa_contratar"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_tarifa_contratar"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_tarifa_contratar"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tension_contratar"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tension_contratar"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_tension_contratar"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_tension_contratar"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_tension"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_tension"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_tp_tension"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_tp_tension"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_crto_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_crto_ps"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_ver_crto_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_ver_crto_ps"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_crto_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_crto_ps"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_estado_crto_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_estado_crto_ps"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_uso_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_uso_ps"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_nom_gestor"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_nom_gestor"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_manual"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_manual"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_medida_baja"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_medida_baja"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_pot_transform"] != System.DBNull.Value)
                            sb.Append(r["nm_pot_transform"].ToString()).Append(",");
                        else
                            sb.Append("null,");

                        if (r["nm_porcentaje_perd"] != System.DBNull.Value)
                            sb.Append(r["nm_porcentaje_perd"].ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["md_fh_efecto_solic"] != System.DBNull.Value)
                            sb.Append("'").Append(r["md_fh_efecto_solic"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["md_fh_efecto_rec"] != System.DBNull.Value)
                            sb.Append("'").Append(r["md_fh_efecto_rec"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_efecto_solic"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_efecto_solic"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_efecto_rec"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_efecto_rec"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_corte_crto"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_corte_crto"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_reenganche_crto"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_reenganche_crto"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_total_rechazos"] != System.DBNull.Value)
                            sb.Append("'").Append(r["nm_total_rechazos"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        for(int i = 1; i <= 10; i++)
                        {
                            if (r["cd_mot_rechazo_" + i] != System.DBNull.Value)
                                sb.Append("'").Append(r["cd_mot_rechazo_" + i].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["de_mot_rechazo_" + i] != System.DBNull.Value)
                                sb.Append("'").Append(r["de_mot_rechazo_" + i].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["de_comentario_mot_rechazo" + i] != System.DBNull.Value)
                                sb.Append("'").Append(r["de_comentario_mot_rechazo" + i].ToString()).Append("',");
                            else
                                sb.Append("null,");
                        }




                        if (r["cd_mot_rechazo_1"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_mot_rechazo_1"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_mot_rechazo_1"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_mot_rechazo_1"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_comentario_mot_rechazo1"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_comentario_mot_rechazo1"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_mot_rechazo_2"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_mot_rechazo_2"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_mot_rechazo_2"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_mot_rechazo_2"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_comentario_mot_rechazo2"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_comentario_mot_rechazo2"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_mot_rechazo_3"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_mot_rechazo_3"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_mot_rechazo_3"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_mot_rechazo_3"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_comentario_mot_rechazo3"] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_comentario_mot_rechazo3"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_mot_rechazo_4"] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r[""] != System.DBNull.Value)
                            sb.Append("'").Append(r[""].ToString()).Append("',");
                        else
                            sb.Append("null,");







                        #endregion

                        if (j == 50)
                        {
                            Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                            firstOnly = true;
                            dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
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
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                        commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        j = 0;
                    }
                }

            }
            catch(Exception ex)
            {

            }
        }

        private string Consulta()
        {
            string strSql = "";

            strSql = "SELECT id, cd_pais, cd_empr, de_empr, cd_cups, id_solicitud_atr, req_name, cd_estado_sol, de_estado_sol, cd_tipo_sol,"
                + " de_tipo_sol, cd_sub_tipo_sol, de_sub_tipo_sol, cd_tipo_motivo_baja, de_tipo_motivo_baja, cd_motivo_denuncia," 
                + " de_motivo_denuncia, fh_creacion, fh_envio, fh_real_envio, fh_aceptacion, fh_rechazo, fh_serv_crto, cd_nif_cif_req," 
                + " cd_tp_cli, cd_nif_cif_cli, tx_nom_cli, tx_apell_cli, de_seg_mercado, cd_linea_negocio, de_dir_cups, cd_empr_distdora," 
                + " de_empr_distdora, name_empr_distdora, cd_crto_atr, de_comentario_sol1, cd_mot_anulacion, de_mot_anulacion, cd_mot_baja," 
                + " de_mot_baja, nm_pot_ctatada_1, nm_pot_ctatada_2, nm_pot_ctatada_3, nm_pot_ctatada_4, nm_pot_ctatada_5, nm_pot_ctatada_6,"
                + " cd_tarifa_contratar, de_tarifa_contratar, cd_tension_contratar, de_tension_contratar, cd_tp_tension, de_tp_tension," 
                + " cd_crto_ps, cd_ver_crto_ps, cd_tp_crto_ps, cd_estado_crto_ps, cd_uso_ps, de_nom_gestor, lg_manual, lg_medida_baja," 
                + " nm_pot_transform, nm_porcentaje_perd, md_fh_efecto_solic, md_fh_efecto_rec, fh_efecto_solic, fh_efecto_rec, lg_corte_crto," 
                + " lg_reenganche_crto, nm_total_rechazos, cd_mot_rechazo_1, de_mot_rechazo_1, de_comentario_mot_rechazo1, cd_mot_rechazo_2," 
                + " de_mot_rechazo_2, de_comentario_mot_rechazo2, cd_mot_rechazo_3, de_mot_rechazo_3, de_comentario_mot_rechazo3, cd_mot_rechazo_4," 
                + " de_mot_rechazo_4, de_comentario_mot_rechazo4, cd_mot_rechazo_5, de_mot_rechazo_5, de_comentario_mot_rechazo5, cd_mot_rechazo_6," 
                + " de_mot_rechazo_6, de_comentario_mot_rechazo6, cd_mot_rechazo_7, de_mot_rechazo_7, de_comentario_mot_rechazo7, cd_mot_rechazo_8," 
                + " de_mot_rechazo_8, de_comentario_mot_rechazo8, cd_mot_rechazo_9, de_mot_rechazo_9, de_comentario_mot_rechazo9, cd_mot_rechazo_10," 
                + " de_mot_rechazo_10, de_comentario_mot_rechazo10, tx_datos_personales, md_fh_solicitud, md_fh_respuesta, fh_respuesta, cd_tarifa," 
                + " de_tarifa, cd_tp_modificacion, de_tp_modificacion, lg_modifyxml, lg_noenvdistribuidora, cd_tipo_sol_adm, de_tipo_sol_adm, cd_cnae, de_cnae," 
                + " fh_env_sistema, fh_envio_anulacion, fh_aceptacion_anulacion, fh_rechazo_anulacion, cd_tipo_crto, de_tipo_crto, cd_tipo_crto_contratar," 
                + " de_tipo_crto_contratar, atr_gestionado_cliente, fh_real_recep_aceptacion, cd_potencia_cont, cd_potencia_cont_2, cd_ciclo, cd_periodos,"
                + " nm_fases, de_tipo_instalacion, lg_cliente_prioritario, cd_cae_actual, cd_cae_cliente, de_territory, de_canal_entrada, cd_de_reclamacion_sve," 
                + " cd_estado_reclamacion, fh_real_recep_rechazo, de_tipo_rechazo, lg_rechazo_tras_aceptacion, lg_rechazo_manual, fh_real_recep_activacion," 
                + " fh_real_recep_acep_anul, fh_real_recep_rech_anul, cd_tension, de_tension, fh_finalizacion_contratar, fh_finalizacion, fh_act_dmco," 
                + " fh_real_env_anul, fh_anul, cd_motivo_cb_atr, de_motivo_cb_atr, id_documento"
                + " FROM ed_owner.t_ed_h_sol_atr where"
                + " cd_pais = 'ESPAÑA'" 
                + " and cd_tp_tension = 'SE'"
                + " and cd_linea_negocio = 'Electricidad'"
                + " and rtrim(ltrim(cd_estado_sol)) not in ('1', '2', '5', '6')"
                + " and cd_sub_tipo_sol <> 'MI'";

            return strSql;
        }
    }
}
