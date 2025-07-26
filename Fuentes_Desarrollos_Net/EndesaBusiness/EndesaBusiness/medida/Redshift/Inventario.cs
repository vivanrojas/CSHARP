using EndesaBusiness.contratacion;
using EndesaBusiness.facturacion;
using EndesaBusiness.punto_suministro;
using EndesaBusiness.servidores;
using EndesaEntity.cnmc.V21_2019_12_17;
using EndesaEntity.contratacion.gas;
using EndesaEntity.facturacion.cuadroDeMando;
using EndesaEntity.medida;
using EndesaEntity;
using iTextSharp.text;
using MySql.Data.MySqlClient;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;

namespace EndesaBusiness.medida.Redshift
{
    public class Inventario
    {
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;
        logs.Log ficheroLog;
        EndesaBusiness.utilidades.Param param;
        public Inventario()
        {
            param = new EndesaBusiness.utilidades.Param("t_ed_h_gest_diar_ps_b2b_param", MySQLDB.Esquemas.MED);
            ss_pp = new utilidades.Seguimiento_Procesos();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_Copia_Inventario_BI");
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
            int num_registros_replace = 50;

            DateTime ultima_actualizacion = new DateTime();

            try
            {

                ultima_actualizacion = Convert.ToDateTime(param.GetValue("ultima_actualizacion"));
                num_registros_replace = Convert.ToInt32(param.GetValue("num_registros_replace"));

                if (param.GetValue("estado_proceso") != "en ejecución" && (ultima_actualizacion.Date < DateTime.Now.Date || param.GetValue("estado_proceso") == "lanzar"))
                {
                    param.UpdateParameter("estado_proceso", "en ejecución");

                    ss_pp.Update_Fecha_Inicio("Medida", "Copia Inventario BI", "Copia Inventario BI");
                    
                    strSql = "delete from t_ed_h_gest_diar_ps_b2b_tmp";
                    ficheroLog.Add(strSql);
                    dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                    commandmy = new MySqlCommand(strSql, dbmy.con);
                    commandmy.ExecuteNonQuery();
                    dbmy.CloseConnection();

                    ultima_actualizacion = UltimaActualizacion();

                    ficheroLog.Add(TotalConsulta(ultima_actualizacion));
                    db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                    command = new OdbcCommand(TotalConsulta(ultima_actualizacion), db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        totalRegistros = Convert.ToInt32(r["total"]);
                    }
                    //db.CloseConnection();


                    ficheroLog.Add(Consulta(ultima_actualizacion));
                    //db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                    command = new OdbcCommand(Consulta(ultima_actualizacion), db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {

                        j++;
                        k++;

                        if (firstOnly)
                        {
                            sb = null;
                            sb = new StringBuilder();
                            sb.Append("REPLACE INTO t_ed_h_gest_diar_ps_b2b_tmp");
                            sb.Append(" (fh_ejecucion, secuencial_registro_entidad_ps, fecha_registro_entidad_ps, cups20, cupspf, segmento_portugal, cliente,");
                            sb.Append(" cif_nif, linea_de_negocio, estado_contrato, tarifa, calendario_distribuidora, calendario_comercializadora, empresa_titular,");
                            sb.Append(" distribuidora, tipo_distribuidora, pais, territorio, potencia_maxima_contratada, tipo_de_pm_calculado, tipo_de_pm_sf, ritmo_facturacion,");
                            sb.Append(" frecuencia_facturacion, prelacion, numero_de_pm_ficticios, numero_de_pm_principales, numero_de_pm_redundantes, tm_ok, codigo_contrato_ps,");
                            sb.Append(" tipo_de_contrato, segmento_front, segmento_back, fecha_de_alta_del_contrato_ps, fecha_de_baja_del_contrato_ps, version_del_contrato_ps,");
                            sb.Append(" fecha_de_inicio_de_version_del_contrato_ps, fecha_de_fin_de_version_del_contrato_ps, codigo_contrato_atr, gestion_propia_atr, revendedor,");
                            sb.Append(" autoconsumo, calculado_con_formula, pendiente_antiguo, suministro_con_perdidas, porcentaje_perdidas, potencia_trafo_perdidas, multipunto,");
                            sb.Append(" producto_horario, agora, retroactividad, medida_aportada_por_cliente, existe_reclamacion, existe_anomalia_segunda_recepcion,");
                            sb.Append(" numero_expediente_abierto, estado_pnt, subestado_pnt, facturacion_agrupada, dia_agrupacion, codigo_agrupacion_cuenta_cliente,");
                            sb.Append(" ranking_tam, ranking_tam_deuda, tam_por_cups, tam_deuda, tam_cif, lg_top, priorizacion_top, tipo_top, anexion_a_top, mera, proceso_concursal,");
                            sb.Append(" suministro_con_riesgo_de_impago, dia_fact_media, dia_fact_max, primer_mes_pendiente_medida, estado_pendiente_medida, fecha_desde_periodo_primer_mes_pte_medida,");
                            sb.Append(" fecha_hasta_periodo_primer_mes_pte_medida, fecha_desde_periodo_primera_orden_de_lectura_mes_pte, fecha_hasta_periodo_primera_orden_de_lectura_mes_pte,");
                            sb.Append(" importe_pendiente, grado_completitud, comentario_primer_mes_pendiente, comentario_ltp_general, subtipo_comentario_ltp_general, usuario_resolucion, grupo_resolucion,");
                            sb.Append(" priorizacion_resolucion, anomalia_primer_mes_pendiente, territorial, nombre_gestor_territorial, responsable_territorial, responsable_zona, gestor_kam,");
                            sb.Append(" tecnologia_via_principal, telefono_via_principal, telefono_via_alternativa, estado_telefono_via_principal, ip_via_principal, puerto_ip_via_principal,");
                            sb.Append(" velocidad_trans_via_principal, formato_trans_via_principal, direccion_de_enlace, direccion_punto_de_medida, clave_lectura, programacion_starbeat,");
                            sb.Append(" indicador_inhibicion, existe_anomalia_puntos_medida, existe_anomalia_telemedida, existe_anomalia_perdidas, existe_anomalia_replica, existe_anomalia_medida,");
                            sb.Append(" existe_anomalia_general, existe_expediente, completitud_formato_f_f011, fecha_maxima_formato_f_f011, completitud_formato_q, fecha_maxima_formato_q, completitud_ree,");
                            sb.Append(" fecha_maxima_ree, completitud_ftp_exabeat_p1d, fecha_maxima_ftp_exabeat_p1d, completitud_ftp_exabeat_p2d, fecha_maxima_ftp_exabeat_p2d, completitud_ftp_exabeat_resumen_atr,");
                            sb.Append(" fecha_maxima_ftp_exabeat_resumen_atr, completitud_ftp_exabeat_f1, fecha_maxima_ftp_exabeat_f1, completitud_starbeat_ch, fecha_maxima_starbeat_ch,");
                            sb.Append(" completitud_starbeat_cch, fecha_maxima_starbeat_cch, completitud_starbeat_resumen_atr, fecha_maxima_starbeat_resumen_atr, completitud_edp_diarias_cch,");
                            sb.Append(" fecha_maxima_edp_diarias_cch, completitud_edp_semanales_cch, fecha_maxima_edp_semanales_cch, completitud_edp_mensuales_cch, fecha_maxima_edp_mensuales_cch,");
                            sb.Append(" completitud_edp_mensuales_corregidas_cch, fecha_maxima_edp_mensuales_corregidas_cch, modo_de_lectura, tipo_pm, version_pm, fecha_movimiento_inventario,");
                            sb.Append(" funcion_pm, primer_mes_pendiente_facturacion, estado_pendiente_facturacion, fecha_desde_periodo_primer_mes_pte_facturacion,");
                            sb.Append(" fecha_hasta_periodo_primer_mes_pte_facturacion, fec_act, cod_carga, created_by, created_date) values ");
                            firstOnly = false;
                        }

                        #region Campos

                        if (r["fh_ejecucion"] != System.DBNull.Value)
                            sb.Append("('").Append(Convert.ToDateTime(r["fh_ejecucion"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("(null,");

                        if (r["secuencial_registro_entidad_ps"] != System.DBNull.Value)
                            sb.Append(r["secuencial_registro_entidad_ps"].ToString()).Append(",");
                        else
                        {
                            sb.Append("null,");
                        }

                        if (r["fecha_registro_entidad_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_registro_entidad_ps"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cups20"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cups20"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cupspf"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cupspf"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["segmento_portugal"] != System.DBNull.Value)
                            sb.Append("'").Append(r["segmento_portugal"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cliente"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["cliente"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cif_nif"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cif_nif"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["linea_de_negocio"] != System.DBNull.Value)
                            sb.Append("'").Append(r["linea_de_negocio"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["estado_contrato"] != System.DBNull.Value)
                            sb.Append("'").Append(r["estado_contrato"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["tarifa"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tarifa"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["calendario_distribuidora"] != System.DBNull.Value)
                            sb.Append("'").Append(r["calendario_distribuidora"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["calendario_comercializadora"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["calendario_comercializadora"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["empresa_titular"] != System.DBNull.Value)
                            sb.Append("'").Append(r["empresa_titular"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["distribuidora"] != System.DBNull.Value)
                            //sb.Append("'").Append(r["distribuidora"].ToString()).Append("',");
                            sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(r["distribuidora"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["tipo_distribuidora"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tipo_distribuidora"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["pais"] != System.DBNull.Value)
                            sb.Append("'").Append(r["pais"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["territorio"] != System.DBNull.Value)
                            sb.Append("'").Append(r["territorio"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["potencia_maxima_contratada"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["potencia_maxima_contratada"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["tipo_de_pm_calculado"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tipo_de_pm_calculado"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["tipo_de_pm_sf"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tipo_de_pm_sf"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["ritmo_facturacion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["ritmo_facturacion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["frecuencia_facturacion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["frecuencia_facturacion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["prelacion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["prelacion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["numero_de_pm_ficticios"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["numero_de_pm_ficticios"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["numero_de_pm_principales"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["numero_de_pm_principales"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["numero_de_pm_redundantes"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["numero_de_pm_redundantes"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["tm_ok"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tm_ok"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["codigo_contrato_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(r["codigo_contrato_ps"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["tipo_de_contrato"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tipo_de_contrato"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["segmento_front"] != System.DBNull.Value)
                            sb.Append("'").Append(r["segmento_front"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["segmento_back"] != System.DBNull.Value)
                            sb.Append("'").Append(r["segmento_back"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_de_alta_del_contrato_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_de_alta_del_contrato_ps"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_de_baja_del_contrato_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_de_baja_del_contrato_ps"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["version_del_contrato_ps"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["version_del_contrato_ps"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["fecha_de_inicio_de_version_del_contrato_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_de_inicio_de_version_del_contrato_ps"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_de_fin_de_version_del_contrato_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_de_fin_de_version_del_contrato_ps"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["codigo_contrato_atr"] != System.DBNull.Value)
                            sb.Append("'").Append(r["codigo_contrato_atr"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["gestion_propia_atr"] != System.DBNull.Value)
                            sb.Append("'").Append(r["gestion_propia_atr"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["revendedor"] != System.DBNull.Value)
                            sb.Append("'").Append(r["revendedor"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["autoconsumo"] != System.DBNull.Value)
                            sb.Append("'").Append(r["autoconsumo"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["calculado_con_formula"] != System.DBNull.Value)
                            sb.Append("'").Append(r["calculado_con_formula"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["pendiente_antiguo"] != System.DBNull.Value)
                            sb.Append("'").Append(r["pendiente_antiguo"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["suministro_con_perdidas"] != System.DBNull.Value)
                            sb.Append("'").Append(r["suministro_con_perdidas"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["porcentaje_perdidas"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["porcentaje_perdidas"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["potencia_trafo_perdidas"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["potencia_trafo_perdidas"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["multipunto"] != System.DBNull.Value)
                            sb.Append("'").Append(r["multipunto"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["producto_horario"] != System.DBNull.Value)
                            sb.Append("'").Append(r["producto_horario"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["agora"] != System.DBNull.Value)
                            sb.Append("'").Append(r["agora"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["retroactividad"] != System.DBNull.Value)
                            sb.Append("'").Append(r["retroactividad"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["medida_aportada_por_cliente"] != System.DBNull.Value)
                            sb.Append("'").Append(r["medida_aportada_por_cliente"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["existe_reclamacion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["existe_reclamacion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["existe_anomalia_segunda_recepcion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["existe_anomalia_segunda_recepcion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["numero_expediente_abierto"] != System.DBNull.Value)
                            sb.Append("'").Append(r["numero_expediente_abierto"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["estado_pnt"] != System.DBNull.Value)
                            sb.Append("'").Append(r["estado_pnt"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["subestado_pnt"] != System.DBNull.Value)
                            sb.Append("'").Append(r["subestado_pnt"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["facturacion_agrupada"] != System.DBNull.Value)
                            sb.Append("'").Append(r["facturacion_agrupada"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["dia_agrupacion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["dia_agrupacion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["codigo_agrupacion_cuenta_cliente"] != System.DBNull.Value)
                            sb.Append("'").Append(r["codigo_agrupacion_cuenta_cliente"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["ranking_tam"] != System.DBNull.Value)
                            sb.Append("'").Append(r["ranking_tam"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["ranking_tam_deuda"] != System.DBNull.Value)
                            sb.Append("'").Append(r["ranking_tam_deuda"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["tam_por_cups"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tam_por_cups"].ToString().Replace(",", ".")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["tam_deuda"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tam_deuda"].ToString().Replace(",", ".")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["tam_cif"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tam_cif"].ToString().Replace(",", ".")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_top"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_top"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["priorizacion_top"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["priorizacion_top"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["tipo_top"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tipo_top"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["anexion_a_top"] != System.DBNull.Value)
                            sb.Append("'").Append(r["anexion_a_top"].ToString()).Append("',");
                        else
                            sb.Append("null,");


                        if (r["mera"] != System.DBNull.Value)
                            sb.Append("'").Append(r["mera"].ToString()).Append("',");
                        else
                            sb.Append("null,");


                        if (r["proceso_concursal"] != System.DBNull.Value)
                            sb.Append("'").Append(r["proceso_concursal"].ToString()).Append("',");
                        else
                            sb.Append("null,");


                        if (r["suministro_con_riesgo_de_impago"] != System.DBNull.Value)
                            sb.Append("'").Append(r["suministro_con_riesgo_de_impago"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["dia_fact_media"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["dia_fact_media"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["dia_fact_max"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["dia_fact_max"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["primer_mes_pendiente_medida"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["primer_mes_pendiente_medida"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["estado_pendiente_medida"] != System.DBNull.Value)
                            sb.Append("'").Append(r["estado_pendiente_medida"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_desde_periodo_primer_mes_pte_medida"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_desde_periodo_primer_mes_pte_medida"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_hasta_periodo_primer_mes_pte_medida"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_hasta_periodo_primer_mes_pte_medida"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_desde_periodo_primera_orden_de_lectura_mes_pte"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_desde_periodo_primera_orden_de_lectura_mes_pte"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_hasta_periodo_primera_orden_de_lectura_mes_pte"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_hasta_periodo_primera_orden_de_lectura_mes_pte"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["importe_pendiente"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["importe_pendiente"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["grado_completitud"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["grado_completitud"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["comentario_primer_mes_pendiente"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(r["comentario_primer_mes_pendiente"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["comentario_ltp_general"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(r["comentario_ltp_general"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["subtipo_comentario_ltp_general"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(r["subtipo_comentario_ltp_general"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["usuario_resolucion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["usuario_resolucion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["grupo_resolucion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["grupo_resolucion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["priorizacion_resolucion"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["priorizacion_resolucion"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["anomalia_primer_mes_pendiente"] != System.DBNull.Value)
                            sb.Append("'").Append(r["anomalia_primer_mes_pendiente"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["territorial"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["territorial"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["nombre_gestor_territorial"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nombre_gestor_territorial"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["responsable_territorial"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["responsable_territorial"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["responsable_zona"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["responsable_zona"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["gestor_kam"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["gestor_kam"]).ToString("gestor_kam").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["tecnologia_via_principal"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tecnologia_via_principal"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["telefono_via_principal"] != System.DBNull.Value)
                            sb.Append("'").Append(r["telefono_via_principal"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["telefono_via_alternativa"] != System.DBNull.Value)
                            sb.Append("'").Append(r["telefono_via_alternativa"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["estado_telefono_via_principal"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["estado_telefono_via_principal"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["ip_via_principal"] != System.DBNull.Value)
                            sb.Append("'").Append(r["ip_via_principal"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["puerto_ip_via_principal"] != System.DBNull.Value)
                            sb.Append("'").Append(r["puerto_ip_via_principal"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["velocidad_trans_via_principal"] != System.DBNull.Value)
                            sb.Append("'").Append(r["velocidad_trans_via_principal"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["formato_trans_via_principal"] != System.DBNull.Value)
                            sb.Append("'").Append(r["formato_trans_via_principal"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["direccion_de_enlace"] != System.DBNull.Value)
                            sb.Append("'").Append(r["direccion_de_enlace"].ToString()).Append("',");
                        else
                            sb.Append("null,");


                        if (r["direccion_punto_de_medida"] != System.DBNull.Value)
                            sb.Append("'").Append(r["direccion_punto_de_medida"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["clave_lectura"] != System.DBNull.Value)
                            sb.Append("'").Append(r["clave_lectura"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["programacion_starbeat"] != System.DBNull.Value)
                            sb.Append("'").Append(r["programacion_starbeat"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["indicador_inhibicion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["indicador_inhibicion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["existe_anomalia_puntos_medida"] != System.DBNull.Value)
                            sb.Append("'").Append(r["existe_anomalia_puntos_medida"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["existe_anomalia_telemedida"] != System.DBNull.Value)
                            sb.Append("'").Append(r["existe_anomalia_telemedida"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["existe_anomalia_perdidas"] != System.DBNull.Value)
                            sb.Append("'").Append(r["existe_anomalia_perdidas"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["existe_anomalia_replica"] != System.DBNull.Value)
                            sb.Append("'").Append(r["existe_anomalia_replica"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["existe_anomalia_medida"] != System.DBNull.Value)
                            sb.Append("'").Append(r["existe_anomalia_medida"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["existe_anomalia_general"] != System.DBNull.Value)
                            sb.Append("'").Append(r["existe_anomalia_general"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["existe_expediente"] != System.DBNull.Value)
                            sb.Append("'").Append(r["existe_expediente"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_formato_f_f011"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_formato_f_f011"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_formato_f_f011"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_formato_f_f011"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_formato_q"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_formato_q"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_formato_q"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_formato_q"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_ree"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_ree"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_ree"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_ree"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_ftp_exabeat_p1d"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_ftp_exabeat_p1d"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_ftp_exabeat_p1d"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_ftp_exabeat_p1d"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_ftp_exabeat_p2d"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_ftp_exabeat_p2d"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_ftp_exabeat_p2d"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_ftp_exabeat_p2d"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_ftp_exabeat_resumen_atr"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_ftp_exabeat_resumen_atr"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_ftp_exabeat_resumen_atr"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_ftp_exabeat_resumen_atr"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_ftp_exabeat_f1"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_ftp_exabeat_f1"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_ftp_exabeat_f1"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_ftp_exabeat_f1"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_starbeat_ch"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_starbeat_ch"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_starbeat_ch"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_starbeat_ch"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_starbeat_cch"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_starbeat_cch"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_starbeat_cch"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_starbeat_cch"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_starbeat_resumen_atr"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_starbeat_resumen_atr"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_starbeat_resumen_atr"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_starbeat_resumen_atr"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_edp_diarias_cch"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_edp_diarias_cch"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_edp_diarias_cch"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_edp_diarias_cch"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_edp_semanales_cch"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_edp_semanales_cch"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_edp_semanales_cch"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_edp_semanales_cch"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_edp_mensuales_cch"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_edp_mensuales_cch"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_edp_mensuales_cch"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_edp_mensuales_cch"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["completitud_edp_mensuales_corregidas_cch"] != System.DBNull.Value)
                            sb.Append("'").Append(r["completitud_edp_mensuales_corregidas_cch"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_maxima_edp_mensuales_corregidas_cch"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_maxima_edp_mensuales_corregidas_cch"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["modo_de_lectura"] != System.DBNull.Value)
                            sb.Append("'").Append(r["modo_de_lectura"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["tipo_pm"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["tipo_pm"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["version_pm"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["version_pm"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["fecha_movimiento_inventario"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_movimiento_inventario"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["funcion_pm"] != System.DBNull.Value)
                            sb.Append("'").Append(r["funcion_pm"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["primer_mes_pendiente_facturacion"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["primer_mes_pendiente_facturacion"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["estado_pendiente_facturacion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["estado_pendiente_facturacion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_desde_periodo_primer_mes_pte_facturacion"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_desde_periodo_primer_mes_pte_facturacion"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fecha_hasta_periodo_primer_mes_pte_facturacion"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fecha_hasta_periodo_primer_mes_pte_facturacion"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
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

                        if (j == num_registros_replace)
                        {
                            Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                            firstOnly = true;
                            dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                            commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                            
                            //ficheroLog.Add(sb.ToString().Substring(0, sb.Length - 1));
                            commandmy.ExecuteNonQuery();
                            //Thread.Sleep(1000);
                            
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
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                        //ficheroLog.Add(sb.ToString().Substring(0, sb.Length - 1));
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        j = 0;
                    }

                    hay_error = false;
                    if (!hay_error)
                    {
                        strSql = "replace into t_ed_h_gest_diar_ps_b2b_hist"
                        + " select t.*, NOW() from t_ed_h_gest_diar_ps_b2b t";
                        ficheroLog.Add(strSql);
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        commandmy = new MySqlCommand(strSql, dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();

                        strSql = "delete from t_ed_h_gest_diar_ps_b2b";
                        ficheroLog.Add(strSql);
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        commandmy = new MySqlCommand(strSql, dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();

                        strSql = "replace into t_ed_h_gest_diar_ps_b2b"
                        + " select * from t_ed_h_gest_diar_ps_b2b_tmp t";
                        ficheroLog.Add(strSql);
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        commandmy = new MySqlCommand(strSql, dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();

                        strSql = "delete from t_ed_h_gest_diar_ps_b2b_tmp";
                        ficheroLog.Add(strSql);
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        commandmy = new MySqlCommand(strSql, dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();

                    }

                    
                    if(param.GetValue("estado_proceso") == "lanzar")
                    {
                        
                        EnvioCorreo_LanzamientoManual();

                    }

                    param.UpdateParameter("estado_proceso", "ejecutado");

                    param.UpdateParameter("ultima_actualizacion", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    ss_pp.Update_Fecha_Fin("Medida", "Copia Inventario BI", "Copia Inventario BI");
                }
            }
            catch(Exception ex)
            {
                param.UpdateParameter("estado_proceso", "ejecutado");
                ficheroLog.AddError("CopiaDatos: " + ex.Message);
                Console.WriteLine(ex.Message);
                ss_pp.Update_Comentario("Medida", "Copia Inventario BI", "Copia Inventario BI", ex.Message);
            }
        }

        private DateTime UltimaActualizacion()
        {
            StringBuilder sb = new StringBuilder();
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "select max(fecha_registro_entidad_ps) as max_fecha from ed_owner.t_ed_h_gest_diar_ps_b2b";
            DateTime f = new DateTime();

            ficheroLog.Add(strSql);
            db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
            command = new OdbcCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["max_fecha"] != System.DBNull.Value)
                    f = Convert.ToDateTime(r["max_fecha"]);
            }
            db.CloseConnection();
            return f;
                
        }


        private string Consulta(DateTime f)
        {
            string strSql = "";

            strSql = "SELECT fh_ejecucion, secuencial_registro_entidad_ps, fecha_registro_entidad_ps, cups20, cupspf, segmento_portugal, cliente,"
                + "cif_nif, linea_de_negocio, estado_contrato, tarifa, calendario_distribuidora, calendario_comercializadora, empresa_titular,"
                + "distribuidora, tipo_distribuidora, pais, territorio, potencia_maxima_contratada, tipo_de_pm_calculado, tipo_de_pm_sf, ritmo_facturacion,"
                + "frecuencia_facturacion, prelacion, numero_de_pm_ficticios, numero_de_pm_principales, numero_de_pm_redundantes, tm_ok, codigo_contrato_ps,"
                + "tipo_de_contrato, segmento_front, segmento_back, fecha_de_alta_del_contrato_ps, fecha_de_baja_del_contrato_ps, version_del_contrato_ps,"
                + "fecha_de_inicio_de_version_del_contrato_ps, fecha_de_fin_de_version_del_contrato_ps, codigo_contrato_atr, gestion_propia_atr, revendedor,"
                + "autoconsumo, calculado_con_formula, pendiente_antiguo, suministro_con_perdidas, porcentaje_perdidas, potencia_trafo_perdidas, multipunto,"
                + "producto_horario, agora, retroactividad, medida_aportada_por_cliente, existe_reclamacion, existe_anomalia_segunda_recepcion,"
                + "numero_expediente_abierto, estado_pnt, subestado_pnt, facturacion_agrupada, dia_agrupacion, codigo_agrupacion_cuenta_cliente,"
                + "ranking_tam, ranking_tam_deuda, tam_por_cups, tam_deuda, tam_cif, lg_top, priorizacion_top, tipo_top, anexion_a_top, mera, proceso_concursal,"
                + "suministro_con_riesgo_de_impago, dia_fact_media, dia_fact_max, primer_mes_pendiente_medida, estado_pendiente_medida, fecha_desde_periodo_primer_mes_pte_medida,"
                + "fecha_hasta_periodo_primer_mes_pte_medida, fecha_desde_periodo_primera_orden_de_lectura_mes_pte, fecha_hasta_periodo_primera_orden_de_lectura_mes_pte,"
                + "importe_pendiente, grado_completitud, comentario_primer_mes_pendiente, comentario_ltp_general, subtipo_comentario_ltp_general, usuario_resolucion, grupo_resolucion,"
                + "priorizacion_resolucion, anomalia_primer_mes_pendiente, territorial, nombre_gestor_territorial, responsable_territorial, responsable_zona, gestor_kam,"
                + "tecnologia_via_principal, telefono_via_principal, telefono_via_alternativa, estado_telefono_via_principal, ip_via_principal, puerto_ip_via_principal,"
                + "velocidad_trans_via_principal, formato_trans_via_principal, direccion_de_enlace, direccion_punto_de_medida, clave_lectura, programacion_starbeat,"
                + "indicador_inhibicion, existe_anomalia_puntos_medida, existe_anomalia_telemedida, existe_anomalia_perdidas, existe_anomalia_replica, existe_anomalia_medida,"
                + "existe_anomalia_general, existe_expediente, completitud_formato_f_f011, fecha_maxima_formato_f_f011, completitud_formato_q, fecha_maxima_formato_q, completitud_ree,"
                + "fecha_maxima_ree, completitud_ftp_exabeat_p1d, fecha_maxima_ftp_exabeat_p1d, completitud_ftp_exabeat_p2d, fecha_maxima_ftp_exabeat_p2d, completitud_ftp_exabeat_resumen_atr,"
                + "fecha_maxima_ftp_exabeat_resumen_atr, completitud_ftp_exabeat_f1, fecha_maxima_ftp_exabeat_f1, completitud_starbeat_ch, fecha_maxima_starbeat_ch,"
                + "completitud_starbeat_cch, fecha_maxima_starbeat_cch, completitud_starbeat_resumen_atr, fecha_maxima_starbeat_resumen_atr, completitud_edp_diarias_cch,"
                + "fecha_maxima_edp_diarias_cch, completitud_edp_semanales_cch, fecha_maxima_edp_semanales_cch, completitud_edp_mensuales_cch, fecha_maxima_edp_mensuales_cch,"
                + "completitud_edp_mensuales_corregidas_cch, fecha_maxima_edp_mensuales_corregidas_cch, modo_de_lectura, tipo_pm, version_pm, fecha_movimiento_inventario,"
                + "funcion_pm, primer_mes_pendiente_facturacion, estado_pendiente_facturacion, fecha_desde_periodo_primer_mes_pte_facturacion,"
                + "fecha_hasta_periodo_primer_mes_pte_facturacion, fec_act, cod_carga"
                + " FROM ed_owner.t_ed_h_gest_diar_ps_b2b where fecha_registro_entidad_ps = '" + f.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'";

            return strSql;
        }

        private string TotalConsulta(DateTime f) 
        {
            string strSql = "";

            strSql = "SELECT COUNT(*) as total "
                + " FROM ed_owner.t_ed_h_gest_diar_ps_b2b where fecha_registro_entidad_ps = '" + f.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'";

            return strSql;
        }

        private void EnvioCorreo_LanzamientoManual()
        {
            EndesaBusiness.utilidades.Global global = new utilidades.Global();
            StringBuilder textBody = new StringBuilder();

            try
            {
                string from = param.GetValue("mail_from");
                string to = global.GetMailUser(param.GetLastUpdateBy("estado_proceso"));
                string cc = null;
                string subject = "Ha finalizado la copia de Inventario BI";

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(" Ha finalizado el proceso de copia de Inventario BI. ");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");

                
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();                
                mes.SendMail(from, to, cc, subject, textBody.ToString());
                ficheroLog.Add("Correo enviado desde: " + param.GetValue("mail_from"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }


    }
}
