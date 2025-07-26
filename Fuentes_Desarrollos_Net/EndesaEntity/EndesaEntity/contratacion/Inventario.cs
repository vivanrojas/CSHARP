using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion
{
    public class Inventario
    {
        // Para guardar la estructura de la tabla t_ed_h_ps

        public string id { get; set; }
        public string cd_pais { get; set; }
        public string cd_tp_tension { get; set; }
        public string cd_empr { get; set; }
        public string cd_cups { get; set; }
        public string cd_cups_ext { get; set; }
        public string cups20 { get; set; }
        public string cd_empr_distdora { get; set; }
        public string de_empr_distdora { get; set; }
        public string de_empr_distdora_nombre { get; set; }
        public DateTime fh_prev_inicio_crto { get; set; }
        public DateTime fh_alta_crto { get; set; }
        public DateTime fh_inicio_vers_crto { get; set; }
        public DateTime fh_prev_baja_crto { get; set; }
        public DateTime fh_prev_fin_crto { get; set; }
        public DateTime fh_baja_crto { get; set; }
        public string de_estado_crto { get; set; }
        public string cd_tarifa_c { get; set; }
        public string nm_tension_actual { get; set; }
        public string cd_tp_disc_horaria { get; set; }
        public string de_tp_disc_horaria { get; set; }
        public string cd_tp_crto { get; set; }
        public string cd_crto_ps { get; set; }
        public int nm_vers_crto { get; set; }
        public double nm_pot_ctatada_1 { get; set; }
        public double nm_pot_ctatada_2 { get; set; }
        public double nm_pot_ctatada_3 { get; set; }
        public double nm_pot_ctatada_4 { get; set; }
        public double nm_pot_ctatada_5 { get; set; }
        public double nm_pot_ctatada_6 { get; set; }
        public double nm_consumo_estimado { get; set; }
        public string cd_crto_comercial { get; set; }
        public string de_seg_mercado { get; set; }
        public string cd_tp_cli { get; set; }
        public string de_tp_cli { get; set; }
        public string tx_nombre_cli { get; set; }
        public string cd_nif_cif_cli { get; set; }
        public string tx_apell_cli { get; set; }
        public string cd_tp_factura { get; set; }
        public string cd_tp_pto_medida { get; set; }
        public string lg_tur { get; set; }
        public string de_eq_medida { get; set; }
        public string de_municip { get; set; }
        public string de_prov { get; set; }
        public string cd_prov_mun { get; set; }
        public string cd_finca_dir { get; set; }
        public string lg_telegestion { get; set; }
        public string lg_bte { get; set; }
        public string lg_gestion_propia { get; set; }
        public string cd_tp_autoconsumo { get; set; }
        public string de_perfil_consumo { get; set; }
        public string lg_medida_baja { get; set; }
        public double nm_porcentaje_perd { get; set; }
        public Int64 nm_pot_transform { get; set; }
        public string cd_tarifa { get; set; }
        public string de_tp_via_ps { get; set; }
        public string de_calle_ps { get; set; }
        public string de_num_ps { get; set; }
        public string de_cp_ps { get; set; }
        public string de_tp_telegestion { get; set; }
        public string cnae_crto_name { get; set; }
        public string lg_ctitular { get; set; }
        public string lg_ccortable { get; set; }
        public string cd_tarifa_pt { get; set; }
        public string de_tarifa_pt { get; set; }
        public string cd_ciclo { get; set; }
        public string lg_cliente_prioritario { get; set; }
        public string de_cliente_prioritario { get; set; }
        public string cd_tp_cli_pt { get; set; }
        public string lg_alteracion_rpe { get; set; }
        public string nm_fases { get; set; }
        public string cd_periodos { get; set; }
        public Int64 cod_carga { get; set; }
        public DateTime fh_act_dmco { get; set; }
        public string m_fh_ejecto { get; set; }
        public DateTime fh_efecto { get; set; }
        public bool lg_migrado_sap { get; set; }
        public string created_by { get; set; }
        public DateTime created_date { get; set; }
        public string last_update_by { get; set; }
        public TimeSpan last_update_date { get; set; }
        public bool existe { get; set; }





    }
}
