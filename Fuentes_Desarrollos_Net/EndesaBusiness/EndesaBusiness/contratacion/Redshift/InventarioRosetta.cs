using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion.Redshift
{
    public class InventarioRosetta
    {
        public Dictionary<string, EndesaEntity.contratacion.Inventario> dic { get; set; }

        public InventarioRosetta(List<string> lista_cups20)
        {
            dic = new Dictionary<string, EndesaEntity.contratacion.Inventario>();
            Carga(lista_cups20);
        }

        private void Carga(List<string> lista_cups20)
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";

            db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
            command = new OdbcCommand(Consulta(lista_cups20), db.con);
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

        private string Consulta(List<string> lista_cups_20)
        {
            string strSql = "";
            bool firstOnly = true;

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
                + " cd_cups_ext in ";

            foreach(string p in lista_cups_20)
            {
                if (firstOnly)
                {
                    strSql  += "('" + p + "'";                        
                    firstOnly = false;
                }
                else
                    strSql += ",'" + p + "'";
            }

            strSql += ")";

            return strSql;
        }
    }
}
