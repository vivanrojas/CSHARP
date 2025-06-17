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
    public class Inventario_PT : EndesaEntity.contratacion.Inventario
    {
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;
        logs.Log ficheroLog;
        EndesaBusiness.utilidades.Param param;

        public Dictionary<string, EndesaEntity.contratacion.Inventario> dic { get; set; }

        public Inventario_PT()
        {
            param = new EndesaBusiness.utilidades.Param("t_ed_h_ps_param", MySQLDB.Esquemas.CON);
            ss_pp = new utilidades.Seguimiento_Procesos();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_Copia_PS_AT_ROSETTA");
            dic = new Dictionary<string, EndesaEntity.contratacion.Inventario>();
        }


        public Inventario_PT(List<string> lista_cups_20)
        {
            param = new EndesaBusiness.utilidades.Param("t_ed_h_ps_param", MySQLDB.Esquemas.CON);
            ss_pp = new utilidades.Seguimiento_Procesos();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_Copia_PS_AT_ROSETTA");
            dic = new Dictionary<string, EndesaEntity.contratacion.Inventario>();
            Carga(lista_cups_20, null);
        }

        public Inventario_PT(List<string> lista_cups_20, List<string> lista_tension)
        {
            param = new EndesaBusiness.utilidades.Param("t_ed_h_ps_param", MySQLDB.Esquemas.CON);
            ss_pp = new utilidades.Seguimiento_Procesos();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_Copia_PS_AT_ROSETTA");
            dic = new Dictionary<string, EndesaEntity.contratacion.Inventario>();
            Carga(lista_cups_20, lista_tension);
        }

        private void Carga(List<string> lista_cups_20, List<string> lista_tension)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(ConsultaMySQL(lista_cups_20, lista_tension), db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.contratacion.Inventario c = new EndesaEntity.contratacion.Inventario();
                #region Campos

                if (r["id"] != System.DBNull.Value)
                    c.id = r["id"].ToString();

                if (r["cd_pais"] != System.DBNull.Value)
                    c.cd_pais = r["cd_pais"].ToString();

                if (r["cd_tp_tension"] != System.DBNull.Value)
                    c.cd_tp_tension = r["cd_tp_tension"].ToString();

                if (r["cd_empr"] != System.DBNull.Value)
                    c.cd_empr = r["cd_empr"].ToString();

                if (r["cd_cups"] != System.DBNull.Value)
                    c.cd_cups = r["cd_cups"].ToString();

                if (r["cd_cups_ext"] != System.DBNull.Value)
                    c.cd_cups_ext = r["cd_cups_ext"].ToString();

                if (r["cups20"] != System.DBNull.Value)
                    c.cups20 = r["cups20"].ToString();

                if (r["cd_empr_distdora"] != System.DBNull.Value)
                    c.cd_empr_distdora = r["cd_empr_distdora"].ToString();

                if (r["de_empr_distdora"] != System.DBNull.Value)
                    c.de_empr_distdora = r["de_empr_distdora"].ToString();

                if (r["de_empr_distdora_nombre"] != System.DBNull.Value)
                    c.de_empr_distdora_nombre = r["de_empr_distdora_nombre"].ToString();

                if (r["fh_prev_inicio_crto"] != System.DBNull.Value)
                    c.fh_prev_inicio_crto = Convert.ToDateTime(r["fh_prev_inicio_crto"]);

                if (r["fh_alta_crto"] != System.DBNull.Value)
                    c.fh_alta_crto = Convert.ToDateTime(r["fh_alta_crto"]);

                if (r["fh_inicio_vers_crto"] != System.DBNull.Value)
                    c.fh_inicio_vers_crto = Convert.ToDateTime(r["fh_inicio_vers_crto"]);

                if (r["fh_prev_baja_crto"] != System.DBNull.Value)
                    c.fh_prev_baja_crto = Convert.ToDateTime(r["fh_prev_baja_crto"]);

                if (r["fh_prev_fin_crto"] != System.DBNull.Value)
                    c.fh_prev_fin_crto = Convert.ToDateTime(r["fh_prev_fin_crto"]);

                if (r["fh_baja_crto"] != System.DBNull.Value)
                    c.fh_baja_crto = Convert.ToDateTime(r["fh_baja_crto"]);

                if (r["de_estado_crto"] != System.DBNull.Value)
                    c.de_estado_crto = r["de_estado_crto"].ToString();

                if (r["cd_tarifa_c"] != System.DBNull.Value)
                    c.cd_tarifa_c = r["cd_tarifa_c"].ToString();

                if (r["nm_tension_actual"] != System.DBNull.Value)
                    c.nm_tension_actual = r["nm_tension_actual"].ToString();

                if (r["cd_tp_disc_horaria"] != System.DBNull.Value)
                    c.cd_tp_disc_horaria = r["cd_tp_disc_horaria"].ToString();

                if (r["de_tp_disc_horaria"] != System.DBNull.Value)
                    c.de_tp_disc_horaria = r["de_tp_disc_horaria"].ToString();

                if (r["cd_tp_crto"] != System.DBNull.Value)
                    c.cd_tp_crto = r["cd_tp_crto"].ToString();

                if (r["cd_crto_ps"] != System.DBNull.Value)
                    c.cd_crto_ps = r["cd_crto_ps"].ToString();

                if (r["nm_vers_crto"] != System.DBNull.Value)
                    c.nm_vers_crto = Convert.ToInt32(r["nm_vers_crto"]);

                if (r["nm_pot_ctatada_1"] != System.DBNull.Value)
                    c.nm_pot_ctatada_1 = Convert.ToDouble(r["nm_pot_ctatada_1"]);

                if (r["nm_pot_ctatada_2"] != System.DBNull.Value)
                    c.nm_pot_ctatada_2 = Convert.ToDouble(r["nm_pot_ctatada_2"]);

                if (r["nm_pot_ctatada_3"] != System.DBNull.Value)
                    c.nm_pot_ctatada_3 = Convert.ToDouble(r["nm_pot_ctatada_3"]);

                if (r["nm_pot_ctatada_4"] != System.DBNull.Value)
                    c.nm_pot_ctatada_4 = Convert.ToDouble(r["nm_pot_ctatada_4"]);

                if (r["nm_pot_ctatada_5"] != System.DBNull.Value)
                    c.nm_pot_ctatada_5 = Convert.ToDouble(r["nm_pot_ctatada_5"]);

                if (r["nm_pot_ctatada_6"] != System.DBNull.Value)
                    c.nm_pot_ctatada_6 = Convert.ToDouble(r["nm_pot_ctatada_6"]);

                if (r["nm_consumo_estimado"] != System.DBNull.Value)
                    c.nm_consumo_estimado = Convert.ToDouble(r["nm_consumo_estimado"]);

                if (r["cd_crto_comercial"] != System.DBNull.Value)
                    c.cd_crto_comercial = r["cd_crto_comercial"].ToString();

                if (r["de_seg_mercado"] != System.DBNull.Value)
                    c.de_seg_mercado = r["de_seg_mercado"].ToString();

                if (r["cd_tp_cli"] != System.DBNull.Value)
                    c.cd_tp_cli = r["cd_tp_cli"].ToString();

                if (r["de_tp_cli"] != System.DBNull.Value)
                    c.de_tp_cli = r["de_tp_cli"].ToString();

                if (r["tx_nombre_cli"] != System.DBNull.Value)
                    c.tx_nombre_cli = r["tx_nombre_cli"].ToString();

                if (r["cd_nif_cif_cli"] != System.DBNull.Value)
                    c.cd_nif_cif_cli = r["cd_nif_cif_cli"].ToString();

                if (r["tx_apell_cli"] != System.DBNull.Value)
                    c.tx_apell_cli = r["tx_apell_cli"].ToString();

                if (r["cd_tp_factura"] != System.DBNull.Value)
                    c.cd_tp_factura = r["cd_tp_factura"].ToString();

                if (r["cd_tp_pto_medida"] != System.DBNull.Value)
                    c.cd_tp_pto_medida = r["cd_tp_pto_medida"].ToString();

                if (r["lg_tur"] != System.DBNull.Value)
                    c.lg_tur = r["lg_tur"].ToString();

                if (r["de_eq_medida"] != System.DBNull.Value)
                    c.de_eq_medida = r["de_eq_medida"].ToString();

                if (r["de_municip"] != System.DBNull.Value)
                    c.de_municip = r["de_municip"].ToString();

                if (r["de_prov"] != System.DBNull.Value)
                    c.de_prov = r["de_prov"].ToString();

                if (r["cd_prov_mun"] != System.DBNull.Value)
                    c.cd_prov_mun = r["cd_prov_mun"].ToString();

                if (r["cd_finca_dir"] != System.DBNull.Value)
                    c.lg_telegestion = r["cd_finca_dir"].ToString();

                if (r["lg_telegestion"] != System.DBNull.Value)
                    c.lg_telegestion = r["lg_telegestion"].ToString();

                if (r["lg_bte"] != System.DBNull.Value)
                    c.lg_bte = r["lg_bte"].ToString();

                if (r["lg_gestion_propia"] != System.DBNull.Value)
                    c.lg_gestion_propia = r["lg_gestion_propia"].ToString();

                if (r["cd_tp_autoconsumo"] != System.DBNull.Value)
                    c.cd_tp_autoconsumo = r["cd_tp_autoconsumo"].ToString();

                if (r["de_perfil_consumo"] != System.DBNull.Value)
                    c.de_perfil_consumo = r["de_perfil_consumo"].ToString();

                if (r["lg_medida_baja"] != System.DBNull.Value)
                    c.lg_medida_baja = r["lg_medida_baja"].ToString();

                if (r["nm_porcentaje_perd"] != System.DBNull.Value)
                    c.nm_porcentaje_perd = Convert.ToDouble(r["nm_porcentaje_perd"]);

                if (r["nm_pot_transform"] != System.DBNull.Value)
                    c.nm_pot_transform = Convert.ToInt64(r["nm_pot_transform"]);

                if (r["cd_tarifa"] != System.DBNull.Value)
                    c.cd_tarifa = r["cd_tarifa"].ToString();

                if (r["de_tp_via_ps"] != System.DBNull.Value)
                    c.de_tp_via_ps = r["de_tp_via_ps"].ToString();

                if (r["de_calle_ps"] != System.DBNull.Value)
                    c.de_calle_ps = r["de_calle_ps"].ToString();

                if (r["de_num_ps"] != System.DBNull.Value)
                    c.de_num_ps = r["de_num_ps"].ToString();

                if (r["de_cp_ps"] != System.DBNull.Value)
                    c.de_cp_ps = r["de_cp_ps"].ToString();

                if (r["de_tp_telegestion"] != System.DBNull.Value)
                    c.de_tp_telegestion = r["de_tp_telegestion"].ToString();

                if (r["cnae_crto_name"] != System.DBNull.Value)
                    c.cnae_crto_name = r["cnae_crto_name"].ToString();

                if (r["lg_ctitular"] != System.DBNull.Value)
                    c.lg_ctitular = r["lg_ctitular"].ToString();

                if (r["lg_ccortable"] != System.DBNull.Value)
                    c.lg_ccortable = r["lg_ccortable"].ToString();

                if (r["cd_tarifa_pt"] != System.DBNull.Value)
                    c.cd_tarifa_pt = r["cd_tarifa_pt"].ToString();

                if (r["de_tarifa_pt"] != System.DBNull.Value)
                    c.de_tarifa_pt = r["de_tarifa_pt"].ToString();

                if (r["cd_ciclo"] != System.DBNull.Value)
                    c.cd_ciclo = r["cd_ciclo"].ToString();

                if (r["lg_cliente_prioritario"] != System.DBNull.Value)
                    c.lg_cliente_prioritario = r["lg_cliente_prioritario"].ToString();

                if (r["de_cliente_prioritario"] != System.DBNull.Value)
                    c.de_cliente_prioritario = r["de_cliente_prioritario"].ToString();

                if (r["cd_tp_cli_pt"] != System.DBNull.Value)
                    c.cd_tp_cli_pt = r["cd_tp_cli_pt"].ToString();

                if (r["lg_alteracion_rpe"] != System.DBNull.Value)
                    c.lg_alteracion_rpe = r["lg_alteracion_rpe"].ToString();

                if (r["nm_fases"] != System.DBNull.Value)
                    c.nm_fases = r["nm_fases"].ToString();

                if (r["cd_periodos"] != System.DBNull.Value)
                    c.cd_periodos = r["cd_periodos"].ToString();

                if (r["cod_carga"] != System.DBNull.Value)
                    c.cod_carga = Convert.ToInt64(r["cod_carga"]);

                if (r["fh_act_dmco"] != System.DBNull.Value)
                    c.fh_act_dmco = Convert.ToDateTime(r["fh_act_dmco"]);

                if (r["m_fh_ejecto"] != System.DBNull.Value)
                    c.m_fh_ejecto = r["m_fh_ejecto"].ToString();

                if (r["fh_efecto"] != System.DBNull.Value)
                    c.fh_efecto = Convert.ToDateTime(r["fh_efecto"]);

                if (r["lg_migrado_sap"] != System.DBNull.Value)
                    c.lg_migrado_sap = r["lg_migrado_sap"].ToString().ToUpper() == "S";

                #endregion

                dic.Add(c.cups20, c);

            }
            db.CloseConnection();

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

                ultima_actualizacion = Convert.ToDateTime(param.GetValue("ultima_actualizacion_pt"));

                if (param.GetValue("estado_proceso_pt") != "en ejecución" && (ultima_actualizacion.Date < DateTime.Now.Date || param.GetValue("estado_proceso_pt") == "lanzar"))
                {
                    param.UpdateParameter("estado_proceso_pt", "en ejecución");

                    ss_pp.Update_Fecha_Inicio("Contratación", "Copia Inventario BI PS_PT_ROSETTA",
                        "Copia Inventario BI PS_PT_ROSETTA");

                    strSql = "delete from t_ed_h_ps_pt_tmp";
                    ficheroLog.Add(strSql);
                    dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    commandmy = new MySqlCommand(strSql, dbmy.con);
                    commandmy.ExecuteNonQuery();
                    dbmy.CloseConnection();

                    ultima_actualizacion = UltimaActualizacion();

                    ficheroLog.Add(Consulta());
                    db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                    command = new OdbcCommand(Consulta(), db.con);
                    command.CommandTimeout = 1000;
                    r = command.ExecuteReader();

                    while (r.Read())
                    {

                        j++;
                        k++;

                        if (firstOnly)
                        {
                            sb.Append("REPLACE INTO t_ed_h_ps_pt_tmp");
                            sb.Append(" (id, cd_pais, cd_tp_tension, cd_empr, cd_cups, cd_cups_ext, cups20, cd_empr_distdora,");
                            sb.Append(" de_empr_distdora, de_empr_distdora_nombre, fh_prev_inicio_crto, fh_alta_crto,");
                            sb.Append(" fh_inicio_vers_crto, fh_prev_baja_crto, fh_prev_fin_crto, fh_baja_crto, de_estado_crto,");
                            sb.Append(" cd_tarifa_c, nm_tension_actual, cd_tp_disc_horaria, de_tp_disc_horaria, cd_tp_crto,");
                            sb.Append(" cd_crto_ps, nm_vers_crto, nm_pot_ctatada_1, nm_pot_ctatada_2, nm_pot_ctatada_3,");
                            sb.Append(" nm_pot_ctatada_4, nm_pot_ctatada_5, nm_pot_ctatada_6, nm_consumo_estimado,");
                            sb.Append(" cd_crto_comercial, de_seg_mercado, cd_tp_cli, de_tp_cli, tx_nombre_cli, cd_nif_cif_cli,");
                            sb.Append(" tx_apell_cli, cd_tp_factura, cd_tp_pto_medida, lg_tur, de_eq_medida, de_municip,");
                            sb.Append(" de_prov, cd_prov_mun, cd_finca_dir, lg_telegestion, lg_bte, lg_gestion_propia,");
                            sb.Append(" cd_tp_autoconsumo, de_perfil_consumo, lg_medida_baja, nm_porcentaje_perd,");
                            sb.Append(" nm_pot_transform, cd_tarifa, de_tp_via_ps, de_calle_ps, de_num_ps, de_cp_ps,");
                            sb.Append(" de_tp_telegestion, cnae_crto_name, lg_ctitular, lg_ccortable, cd_tarifa_pt,");
                            sb.Append(" de_tarifa_pt, cd_ciclo, lg_cliente_prioritario, de_cliente_prioritario, cd_tp_cli_pt,");
                            sb.Append(" lg_alteracion_rpe, nm_fases, cd_periodos, cod_carga, fh_act_dmco, m_fh_ejecto,");
                            sb.Append(" fh_efecto, lg_migrado_sap, created_by, created_date) values ");
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

                        if (r["cd_tp_tension"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_tension"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_empr"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_empr"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_cups"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_cups"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_cups_ext"] != System.DBNull.Value)
                        {
                            sb.Append("'").Append(r["cd_cups_ext"].ToString()).Append("',");
                            sb.Append("'").Append(r["cd_cups_ext"].ToString().Substring(0, 20)).Append("',");
                        }
                        else
                            sb.Append("null,null,");

                        if (r["cd_empr_distdora"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_empr_distdora"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_empr_distdora"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_empr_distdora"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_empr_distdora_nombre"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_empr_distdora_nombre"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_prev_inicio_crto"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_prev_inicio_crto"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_alta_crto"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_alta_crto"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_inicio_vers_crto"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_inicio_vers_crto"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_prev_baja_crto"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_prev_baja_crto"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_prev_fin_crto"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_prev_fin_crto"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_baja_crto"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_baja_crto"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_estado_crto"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_estado_crto"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tarifa_c"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tarifa_c"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_tension_actual"] != System.DBNull.Value)
                            sb.Append("'").Append(r["nm_tension_actual"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_disc_horaria"] != System.DBNull.Value)
                        {
                            sb.Append("'").Append(r["cd_tp_disc_horaria"].ToString()).Append("',");
                        }
                        else
                            sb.Append("null,");

                        if (r["de_tp_disc_horaria"] != System.DBNull.Value)
                        {
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_tp_disc_horaria"].ToString())).Append("',");
                        }
                        else
                            sb.Append("null,");

                        if (r["cd_tp_crto"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_crto"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_crto_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_crto_ps"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_vers_crto"] != System.DBNull.Value)
                            sb.Append(r["nm_vers_crto"].ToString()).Append(",");
                        else
                            sb.Append("null,");

                        for (int i = 1; i <= 6; i++)
                        {
                            if (r["nm_pot_ctatada_" + i] != System.DBNull.Value)
                                sb.Append(r["nm_pot_ctatada_" + i].ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");
                        }

                        if (r["nm_consumo_estimado"] != System.DBNull.Value)
                            sb.Append(r["nm_consumo_estimado"].ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["cd_crto_comercial"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_crto_comercial"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_seg_mercado"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_seg_mercado"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_cli"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_cli"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_tp_cli"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_tp_cli"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["tx_nombre_cli"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["tx_nombre_cli"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_nif_cif_cli"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_nif_cif_cli"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["tx_apell_cli"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["tx_apell_cli"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_factura"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_factura"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_pto_medida"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["cd_tp_pto_medida"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_tur"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_tur"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_eq_medida"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_eq_medida"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_municip"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_municip"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_prov"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_prov"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_prov_mun"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_prov_mun"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_finca_dir"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_finca_dir"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_telegestion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_telegestion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_bte"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_bte"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_gestion_propia"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_gestion_propia"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_autoconsumo"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_autoconsumo"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_perfil_consumo"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_perfil_consumo"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_medida_baja"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_medida_baja"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_porcentaje_perd"] != System.DBNull.Value)
                            sb.Append(r["nm_porcentaje_perd"].ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["nm_pot_transform"] != System.DBNull.Value)
                            sb.Append(r["nm_pot_transform"].ToString()).Append(",");
                        else
                            sb.Append("null,");

                        if (r["cd_tarifa"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tarifa"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_tp_via_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_tp_via_ps"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_calle_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_calle_ps"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_num_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_num_ps"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_cp_ps"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_cp_ps"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_tp_telegestion"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_tp_telegestion"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cnae_crto_name"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["cnae_crto_name"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_ctitular"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_ctitular"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_ccortable"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_ccortable"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tarifa_pt"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tarifa_pt"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_tarifa_pt"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_tarifa_pt"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_ciclo"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["cd_ciclo"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_cliente_prioritario"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_cliente_prioritario"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_cliente_prioritario"] != System.DBNull.Value)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(r["de_cliente_prioritario"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_cli_pt"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_cli_pt"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_alteracion_rpe"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_alteracion_rpe"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_fases"] != System.DBNull.Value)
                            sb.Append("'").Append(r["nm_fases"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_periodos"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_periodos"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cod_carga"] != System.DBNull.Value)
                            sb.Append(r["cod_carga"].ToString()).Append(",");
                        else
                            sb.Append("null,");

                        if (r["fh_act_dmco"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_act_dmco"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["m_fh_ejecto"] != System.DBNull.Value)
                            sb.Append("'").Append(r["m_fh_ejecto"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_efecto"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_efecto"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_migrado_sap"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_migrado_sap"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                        sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                        #endregion

                        if (j == 150)
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

                    hay_error = false;
                    if (!hay_error)
                    {
                        strSql = "replace into t_ed_h_ps_pt_hist"
                        + " select t.* from t_ed_h_ps_pt t";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                        commandmy = new MySqlCommand(strSql, dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();

                        strSql = "delete from t_ed_h_ps_pt";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                        commandmy = new MySqlCommand(strSql, dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();

                        strSql = "replace into t_ed_h_ps_pt"
                        + " select * from t_ed_h_ps_pt_tmp t";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                        commandmy = new MySqlCommand(strSql, dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();

                        strSql = "delete from t_ed_h_ps_pt_tmp";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                        commandmy = new MySqlCommand(strSql, dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();

                    }

                    param.UpdateParameter("estado_proceso_pt", "ejecutado");
                    param.UpdateParameter("ultima_actualizacion_pt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                    ss_pp.Update_Fecha_Fin("Contratación", "Copia Inventario BI PS_PT_ROSETTA",
                        "Copia Inventario BI PS_PT_ROSETTA");

                }
            }
            catch (Exception ex)
            {
                param.UpdateParameter("estado_proceso", "ejecutado");
                ficheroLog.addError("CopiaDatos: " + ex.Message);
                Console.WriteLine(ex.Message);
                ss_pp.Update_Comentario("Contratación", "Copia Inventario BI PS_PT_ROSETTA", "Copia Inventario BI PS_PT_ROSETTA", ex.Message);
            }
        }

        private DateTime UltimaActualizacion()
        {
            StringBuilder sb = new StringBuilder();
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "select max(fh_act_dmco) as max_fecha from ed_owner.t_ed_h_ps where cd_pais = 'PORTUGAL'";
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

        private string Consulta()
        {
            string strSql = "";

            strSql = "SELECT id, cd_pais, cd_tp_tension, cd_empr, cd_cups, cd_cups_ext, cd_empr_distdora,"
                + " de_empr_distdora, de_empr_distdora_nombre, fh_prev_inicio_crto, fh_alta_crto,"
                + " fh_inicio_vers_crto, fh_prev_baja_crto, fh_prev_fin_crto, fh_baja_crto, de_estado_crto,"
                + " cd_tarifa_c, nm_tension_actual, cd_tp_disc_horaria, de_tp_disc_horaria, cd_tp_crto,"
                + " cd_crto_ps, nm_vers_crto, nm_pot_ctatada_1, nm_pot_ctatada_2, nm_pot_ctatada_3, nm_pot_ctatada_4,"
                + " nm_pot_ctatada_5, nm_pot_ctatada_6, nm_consumo_estimado, cd_crto_comercial, de_seg_mercado,"
                + " cd_tp_cli, de_tp_cli, tx_nombre_cli, cd_nif_cif_cli, tx_apell_cli, cd_tp_factura, cd_tp_pto_medida,"
                + " lg_tur, de_eq_medida, de_municip, de_prov, cd_prov_mun, cd_finca_dir, lg_telegestion, lg_bte,"
                + " lg_gestion_propia, cd_tp_autoconsumo, de_perfil_consumo, lg_medida_baja, nm_porcentaje_perd,"
                + " nm_pot_transform, cd_tarifa, de_tp_via_ps, de_calle_ps, de_num_ps, de_cp_ps, de_tp_telegestion,"
                + " cnae_crto_name, lg_ctitular, lg_ccortable, cd_tarifa_pt, de_tarifa_pt, cd_ciclo,"
                + " lg_cliente_prioritario, de_cliente_prioritario, cd_tp_cli_pt, lg_alteracion_rpe, nm_fases,"
                + " cd_periodos, cod_carga, fh_act_dmco, m_fh_ejecto, fh_efecto, lg_migrado_sap"
                + " FROM ed_owner.t_ed_h_ps where"
                + " cd_pais = 'PORTUGAL'";

            return strSql;
        }

        private string ConsultaMySQL(List<string> lista_cups_20, List<string> lista_tension)
        {
            StringBuilder sb = new StringBuilder();

            if (lista_cups_20 != null && lista_tension != null)
            {
                sb.Append("SELECT id, cd_pais, cd_tp_tension, cd_empr, cd_cups, cd_cups_ext, cups20, cd_empr_distdora,");
                sb.Append(" de_empr_distdora, de_empr_distdora_nombre, fh_prev_inicio_crto, fh_alta_crto,");
                sb.Append(" fh_inicio_vers_crto, fh_prev_baja_crto, fh_prev_fin_crto, fh_baja_crto, de_estado_crto,");
                sb.Append(" cd_tarifa_c, nm_tension_actual, cd_tp_disc_horaria, de_tp_disc_horaria, cd_tp_crto,");
                sb.Append(" cd_crto_ps, nm_vers_crto, nm_pot_ctatada_1, nm_pot_ctatada_2, nm_pot_ctatada_3, nm_pot_ctatada_4,");
                sb.Append(" nm_pot_ctatada_5, nm_pot_ctatada_6, nm_consumo_estimado, cd_crto_comercial, de_seg_mercado,");
                sb.Append(" cd_tp_cli, de_tp_cli, tx_nombre_cli, cd_nif_cif_cli, tx_apell_cli, cd_tp_factura, cd_tp_pto_medida,");
                sb.Append(" lg_tur, de_eq_medida, de_municip, de_prov, cd_prov_mun, cd_finca_dir, lg_telegestion, lg_bte,");
                sb.Append(" lg_gestion_propia, cd_tp_autoconsumo, de_perfil_consumo, lg_medida_baja, nm_porcentaje_perd,");
                sb.Append(" nm_pot_transform, cd_tarifa, de_tp_via_ps, de_calle_ps, de_num_ps, de_cp_ps, de_tp_telegestion,");
                sb.Append(" cnae_crto_name, lg_ctitular, lg_ccortable, cd_tarifa_pt, de_tarifa_pt, cd_ciclo,");
                sb.Append(" lg_cliente_prioritario, de_cliente_prioritario, cd_tp_cli_pt, lg_alteracion_rpe, nm_fases,");
                sb.Append(" cd_periodos, cod_carga, fh_act_dmco, m_fh_ejecto, fh_efecto, lg_migrado_sap");
                sb.Append(" FROM t_ed_h_ps_pt where");
                sb.Append(" cups20 in ('").Append(lista_cups_20[0]).Append("'");
                for (int i = 1; i < lista_cups_20.Count; i++)
                    sb.Append(", '" + lista_cups_20[i] + "'");
                sb.Append(")");

                if(lista_tension.Count > 0)
                {
                    sb.Append(" and cd_tp_tension in ('").Append(lista_tension[0]).Append("'");
                    for (int i = 1; i < lista_tension.Count; i++)
                        sb.Append(", '" + lista_tension[i] + "'");
                    sb.Append(")");
                }

            }
            else if (lista_tension != null)
            {
                sb.Append("SELECT id, cd_pais, cd_tp_tension, cd_empr, cd_cups, cd_cups_ext, cups20, cd_empr_distdora,");
                sb.Append(" de_empr_distdora, de_empr_distdora_nombre, fh_prev_inicio_crto, fh_alta_crto,");
                sb.Append(" fh_inicio_vers_crto, fh_prev_baja_crto, fh_prev_fin_crto, fh_baja_crto, de_estado_crto,");
                sb.Append(" cd_tarifa_c, nm_tension_actual, cd_tp_disc_horaria, de_tp_disc_horaria, cd_tp_crto,");
                sb.Append(" cd_crto_ps, nm_vers_crto, nm_pot_ctatada_1, nm_pot_ctatada_2, nm_pot_ctatada_3, nm_pot_ctatada_4,");
                sb.Append(" nm_pot_ctatada_5, nm_pot_ctatada_6, nm_consumo_estimado, cd_crto_comercial, de_seg_mercado,");
                sb.Append(" cd_tp_cli, de_tp_cli, tx_nombre_cli, cd_nif_cif_cli, tx_apell_cli, cd_tp_factura, cd_tp_pto_medida,");
                sb.Append(" lg_tur, de_eq_medida, de_municip, de_prov, cd_prov_mun, cd_finca_dir, lg_telegestion, lg_bte,");
                sb.Append(" lg_gestion_propia, cd_tp_autoconsumo, de_perfil_consumo, lg_medida_baja, nm_porcentaje_perd,");
                sb.Append(" nm_pot_transform, cd_tarifa, de_tp_via_ps, de_calle_ps, de_num_ps, de_cp_ps, de_tp_telegestion,");
                sb.Append(" cnae_crto_name, lg_ctitular, lg_ccortable, cd_tarifa_pt, de_tarifa_pt, cd_ciclo,");
                sb.Append(" lg_cliente_prioritario, de_cliente_prioritario, cd_tp_cli_pt, lg_alteracion_rpe, nm_fases,");
                sb.Append(" cd_periodos, cod_carga, fh_act_dmco, m_fh_ejecto, fh_efecto, lg_migrado_sap");
                sb.Append(" FROM t_ed_h_ps_pt where");                
                sb.Append(" cd_tp_tension in ('").Append(lista_tension[0]).Append("'");
                for (int i = 1; i < lista_tension.Count; i++)
                    sb.Append(", '" + lista_tension[i] + "'");
                sb.Append(")");
                
            }
            else
            {
                sb.Append("SELECT id, cd_pais, cd_tp_tension, cd_empr, cd_cups, cd_cups_ext, cups20, cd_empr_distdora,");
                sb.Append(" de_empr_distdora, de_empr_distdora_nombre, fh_prev_inicio_crto, fh_alta_crto,");
                sb.Append(" fh_inicio_vers_crto, fh_prev_baja_crto, fh_prev_fin_crto, fh_baja_crto, de_estado_crto,");
                sb.Append(" cd_tarifa_c, nm_tension_actual, cd_tp_disc_horaria, de_tp_disc_horaria, cd_tp_crto,");
                sb.Append(" cd_crto_ps, nm_vers_crto, nm_pot_ctatada_1, nm_pot_ctatada_2, nm_pot_ctatada_3, nm_pot_ctatada_4,");
                sb.Append(" nm_pot_ctatada_5, nm_pot_ctatada_6, nm_consumo_estimado, cd_crto_comercial, de_seg_mercado,");
                sb.Append(" cd_tp_cli, de_tp_cli, tx_nombre_cli, cd_nif_cif_cli, tx_apell_cli, cd_tp_factura, cd_tp_pto_medida,");
                sb.Append(" lg_tur, de_eq_medida, de_municip, de_prov, cd_prov_mun, cd_finca_dir, lg_telegestion, lg_bte,");
                sb.Append(" lg_gestion_propia, cd_tp_autoconsumo, de_perfil_consumo, lg_medida_baja, nm_porcentaje_perd,");
                sb.Append(" nm_pot_transform, cd_tarifa, de_tp_via_ps, de_calle_ps, de_num_ps, de_cp_ps, de_tp_telegestion,");
                sb.Append(" cnae_crto_name, lg_ctitular, lg_ccortable, cd_tarifa_pt, de_tarifa_pt, cd_ciclo,");
                sb.Append(" lg_cliente_prioritario, de_cliente_prioritario, cd_tp_cli_pt, lg_alteracion_rpe, nm_fases,");
                sb.Append(" cd_periodos, cod_carga, fh_act_dmco, m_fh_ejecto, fh_efecto, lg_migrado_sap");
                sb.Append(" FROM t_ed_h_ps_pt");

            }

            return sb.ToString();
        }

        public bool ExisteCUPS(string cups20)
        {            
            EndesaEntity.contratacion.Inventario o;
            return (dic.TryGetValue(cups20, out o));            
        }
        public bool ExisteCIF(string cif)
        {
            bool existe = false;
            EndesaEntity.contratacion.PS_AT_Tabla o;
            foreach (KeyValuePair<string, EndesaEntity.contratacion.Inventario> p in dic)
            {
                if (p.Value.cd_nif_cif_cli == cif)
                {
                    existe = true;
                    break;
                }
            }

            return existe;
        }

        public void GetCUPS(string cups20)
        {
            EndesaEntity.contratacion.Inventario o;
            if(dic.TryGetValue(cups20, out o))
            {
                this.existe = true;
                this.lg_migrado_sap = o.lg_migrado_sap;

            }
            else
            {
                this.existe = false;
            }
        }


    }
}
