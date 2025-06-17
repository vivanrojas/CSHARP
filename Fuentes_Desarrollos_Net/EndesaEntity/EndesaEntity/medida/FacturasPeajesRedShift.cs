using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Instrumentation;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace EndesaEntity.medida
{
    public class FacturasPeajesRedShift
    {
        public string cl_crto_ext_ps { get; set; }
        public string id_crto_ext_ps { get; set; }
        public string id_crto_ps { get; set; }
        public int cd_sec_fact { get; set; }
        public string cd_linea_neg { get; set; }
        public string de_linea_neg { get; set; }
        public string cd_empr_distr { get; set; }
        public string cd_est_fact { get; set; }
        public string cd_mes { get; set; }
        public string id_fact { get; set; }
        public DateTime fh_fact { get; set; }
        public DateTime fh_ini_fact { get; set; }
        public DateTime fh_fin_fact { get; set; }
        public double nm_importe { get; set; }
        public double nm_meses { get; set; }
        public string cd_tp_fact { get; set; }
        public string de_tp_fact { get; set; }
        public string id_fact_sust { get; set; }
        public string cd_cups { get; set; }
        public string cd_cups_ext { get; set; }
        public string cd_empr_distr_cne { get; set;}
        public string de_empr_distr_cne { get; set; }
        public double[] nm_med_potencia { get; set; }
        public double[] nm_prec_potencia { get; set; }
        public double[] nm_med_activa { get; set; }
        public double[] nm_prec_activa { get; set; }
        public double[] nm_med_reactiva { get; set; }
        public double[] nm_prec_reactiva { get; set; }
        public string[] cd_concepto { get; set; }
        public string[] cd_concepto_sce { get; set; }
        public string[] de_concepto { get; set; }
        public double[] im_concepto { get; set; }
        public string[] de_impuesto { get; set; }
        public double[] nm_porcentaje { get; set; }
        public double[] nm_base { get; set; }
        public double[] nm_importe_impuesto { get; set; }
        public int cod_carga_ods { get; set; }
        public DateTime fh_act_ods { get; set; }
        public DateTime fh_act_dmco { get; set; }
        public string cd_tarifa { get; set; }
        public string cd_tarifa_ff { get; set; }        
        public int nm_pot_ctatada { get; set; }
        public int cd_comercializadora { get; set; }
        public string cd_tipo_consumo { get; set; }
        public string cd_municipio { get; set; }
        public string cd_cups20_metra { get; set; }
        public string cd_tp_teleg { get; set; }
        public double consumo_tot_act { get; set; }
        public double consumo_punta { get; set; }
        public double consumo_llano { get; set; }
        public double consumo_valle { get; set; }
        public double consumo_activa4 { get; set; }
        public double consumo_activa5 { get; set; }
        public double consumo_activa6 { get; set; }
        public double consumo_tot_react { get; set; }
        public double[] consumo_reactiva { get; set; }
        public double potencia_maxpunta { get; set; }
        public double potencia_maxllano { get; set; }
        public double potencia_maxvalle { get; set; }
        public double potencia_max4 { get; set; }
        public double potencia_max5 { get; set; }
        public double potencia_max6 { get; set; }
        public double consumo_supervalle { get; set; }
        public DateTime fh_recepcion { get; set; }
        public int cod_carga { get; set; }
        public string id_pseudofactura { get; set; }
        public int cd_comercializadora_fact { get; set; }
        public string cd_empr_comer_cne { get; set; }
        public string id_fact_recti { get; set; }
        public string id_nif_dist { get; set; }
        public string de_comcnmc { get; set; }
        public double im_salfact { get; set; }
        public string tp_moneda { get; set; }
        public DateTime fh_feboe { get; set; }
        public string lg_med_perdida { get; set; }
        public long nm_vastrafo { get; set; }
        public double nm_porc_perdida { get; set; }
        public string cd_fact_curva { get; set; }
        public DateTime fh_desde_cch { get; set; }
        public DateTime fh_hasta_cch { get; set; }
        public double nm_num_mes { get; set; }
        public DateTime fh_lect_ant_punta { get; set; }
        public string cd_tp_fuente_ant_punta { get; set; }
        public double nm_lect_ant_punta { get; set; }
        public DateTime fh_lect_act_punta { get; set; }
        public string cd_tp_fuente_act_punta { get; set; }
        public double nm_lect_act_punta { get; set; }
        public string im_ajuste_integr_punta { get; set; }
        public string cd_tp_anoma_punta { get; set; }
        public DateTime fh_lect_ant_llano { get; set; }
        public string cd_tp_fuente_ant_llano { get; set; }
        public double nm_lect_ant_llano { get; set; }
        public DateTime fh_lect_act_llano { get; set; }
        public string cd_tp_fuente_act_llano { get; set; }
        public double nm_lect_act_llano { get; set; }
        public string im_ajuste_integr_llano { get; set; }
        public string cd_tp_anoma_llano { get; set; }
        public DateTime fh_lect_ant_valle { get; set; }
        public string cd_tp_fuente_ant_valle { get; set; }
        public double nm_lect_ant_valle { get; set; }
        public DateTime fh_lect_act_valle { get; set; }
        public string cd_tp_fuente_act_valle { get; set; }
        public double nm_lect_act_valle { get; set; }
        public string im_ajuste_integr_valle { get; set; }
        public string cd_tp_anoma_valle { get; set; }

        public DateTime fh_lect_ant_activa4 { get; set; }
        public string cd_tp_fuente_ant_activa4 { get; set; }
        public double nm_lect_ant_activa4 { get; set; }
        public DateTime fh_lect_act_activa4 { get; set; }
        public string cd_tp_fuente_act_activa4 { get; set; }
        public double nm_lect_act_activa4 { get; set; }
        public string im_ajuste_integr_activa4 { get; set; }
        public string cd_tp_anoma_activa4 { get; set; }

        public DateTime fh_lect_ant_activa5 { get; set; }
        public string cd_tp_fuente_ant_activa5 { get; set; }
        public double nm_lect_ant_activa5 { get; set; }
        public DateTime fh_lect_act_activa5 { get; set; }
        public string cd_tp_fuente_act_activa5 { get; set; }
        public double nm_lect_act_activa5 { get; set; }
        public string im_ajuste_integr_activa5 { get; set; }
        public string cd_tp_anoma_activa5 { get; set; }

        public DateTime fh_lect_ant_activa6 { get; set; }
        public string cd_tp_fuente_ant_activa6 { get; set; }
        public double nm_lect_ant_activa6 { get; set; }
        public DateTime fh_lect_act_activa6 { get; set; }
        public string cd_tp_fuente_act_activa6 { get; set; }
        public double nm_lect_act_activa6 { get; set; }
        public string im_ajuste_integr_activa6 { get; set; }
        public string cd_tp_anoma_activa6 { get; set; }

        public DateTime[] fh_lect_ant_reactiva { get; set; }
        public string[] cd_tp_fuente_ant_reactiva { get; set; }
        public double[] nm_lect_ant_reactiva { get; set; }
        public DateTime[] fh_lect_act_reactiva { get; set; }
        public string[] cd_tp_fuente_act_reactiva { get; set; }
        public double[] nm_lect_act_reactiva { get; set; }
        public string[] im_ajuste_integr_reactiva { get; set; }
        public string[] cd_tp_anoma_reactiva { get; set; }

        public DateTime fh_lect_ant_maxpunta { get; set; }
        public string cd_tp_fuente_ant_maxpunta { get; set; }
        public double nm_lect_ant_maxpunta { get; set; }
        public DateTime fh_lect_act_maxpunta { get; set; }
        public string cd_tp_fuente_act_maxpunta { get; set; }
        public double nm_lect_act_maxpunta { get; set; }
        public string im_ajuste_integr_maxpunta { get; set; }
        public string cd_tp_anoma_maxpunta { get; set; }
        public DateTime fh_lect_ant_maxllano { get; set; }
        public string cd_tp_fuente_ant_maxllano { get; set; }
        public double nm_lect_ant_maxllano { get; set; }
        public DateTime fh_lect_act_maxllano { get; set; }
        public string cd_tp_fuente_act_maxllano { get; set; }
        public double nm_lect_act_maxllano { get; set; }
        public string im_ajuste_integr_maxllano { get; set; }
        public string cd_tp_anoma_maxllano { get; set; }
        public DateTime fh_lect_ant_maxvalle { get; set; }
        public string cd_tp_fuente_ant_maxvalle { get; set; }
        public double nm_lect_ant_maxvalle { get; set; }
        public DateTime fh_lect_act_maxvalle { get; set; }
        public string cd_tp_fuente_act_maxvalle { get; set; }
        public double nm_lect_act_maxvalle { get; set; }
        public string im_ajuste_integr_maxvalle { get; set; }
        public string cd_tp_anoma_maxvalle { get; set; }

        public DateTime fh_lect_ant_max4 { get; set; }
        public string cd_tp_fuente_ant_max4 { get; set; }
        public double nm_lect_ant_max4 { get; set; }
        public DateTime fh_lect_act_max4 { get; set; }
        public string cd_tp_fuente_act_max4 { get; set; }
        public double nm_lect_act_max4 { get; set; }
        public string im_ajuste_integr_max4 { get; set; }
        public string cd_tp_anoma_max4 { get; set; }

        public DateTime fh_lect_ant_max5 { get; set; }
        public string cd_tp_fuente_ant_max5 { get; set; }
        public double nm_lect_ant_max5 { get; set; }
        public DateTime fh_lect_act_max5 { get; set; }
        public string cd_tp_fuente_act_max5 { get; set; }
        public double nm_lect_act_max5 { get; set; }
        public string im_ajuste_integr_max5 { get; set; }
        public string cd_tp_anoma_max5 { get; set; }

        public DateTime fh_lect_ant_max6 { get; set; }
        public string cd_tp_fuente_ant_max6 { get; set; }
        public double nm_lect_ant_max6 { get; set; }
        public DateTime fh_lect_act_max6 { get; set; }
        public string cd_tp_fuente_act_max6 { get; set; }
        public double nm_lect_act_max6 { get; set; }
        public string im_ajuste_integr_max6 { get; set; }
        public string cd_tp_anoma_max6 { get; set; }

        public DateTime fh_lect_ant_supervalle { get; set; }
        public string cd_tp_fuente_ant_supervalle { get; set; }
        public double nm_lect_ant_supervalle { get; set; }
        public DateTime fh_lect_act_supervalle { get; set; }
        public string cd_tp_fuente_act_supervalle { get; set; }
        public double nm_lect_act_supervalle { get; set; }
        public string im_ajuste_integr_supervalle { get; set; }
        public string cd_tp_anoma_supervalle { get; set; }

        public string[] cd_concepto_reper { get; set; }
        public double[] im_concepto_reper { get; set; }

        public int cd_cmunicip_atr { get; set; }
        public string de_est_fact { get; set; }

        public DateTime fh_desde_ener { get; set; }
        public DateTime fh_hasta_ener { get; set; }
        public DateTime fh_desde_pot { get; set; }
        public DateTime fh_hasta_pot { get; set; }
        public long cd_modo_crtl_pot { get; set; }
        public long nm_dias_fact { get; set; }
        public double im_tot_fact { get; set; }
        public double im_saldo_tot_fact { get; set; }
        public long nm_tot_recibos { get; set; }
        public DateTime fh_valor { get; set; }
        public DateTime fh_lim_pago { get; set; }
        public string id_remesa { get; set; }
        public string cd_munic_red { get; set; }
        public string lg_penaliza_icp { get; set; }
        public string identificador { get; set; }
        public string cd_mensaje { get; set; }
        public string cd_doc_ref { get; set; }
        public string cd_tp_doc { get; set; }
        public string de_tp_doc { get; set; }
        public string de_tp_aceso_serv { get; set; }
        public string lg_relevancia_recl_pdp { get; set; }
        public string de_empr_comer_cne { get; set; }
        public string cd_huella_fact_atr { get; set; }
        public string fichero { get; set; }
        public string cd_huella_blq_fich { get; set; }
        public DateTime fh_publicacion { get; set; }
        public DateTime fh_public_fichero { get; set; }
        public string cd_solicitud { get; set; }
        public string cd_sec_solicitud { get; set; }
        public string lg_atr { get; set; }
        public string cd_tp_reg { get; set; }
        public string de_tp_reg { get; set; }
        public string cd_expdte { get; set; }
        public string lg_bloq_af { get; set; }
        public string lg_recl_abta { get; set; }
        public long cd_origen_fact { get; set; }
        public string de_origen_fact { get; set; }
        public string cd_aniofactura { get; set; }
        public long cd_pais { get; set; }
        public string lg_autoconsumo { get; set; }
        public string cd_tp_pm { get; set; }
        public string lg_duracioninfanio { get; set; }
        public DateTime fh_creacion { get; set; }
        public long cd_usuario_creador { get; set; }
        public string de_usuario_creador { get; set; }
        public DateTime fh_mod { get; set; }
        public long cd_usuario_mod { get; set; }
        public string de_usuario_mod { get; set; }
        public long cd_ubicacion { get; set; }
        public string de_ubicacion { get; set; }
        public string id_exp_af { get; set; }
        public string lg_registro_borrado { get; set; }
        public string de_tipo_borrado { get; set; }
        public string cd_rg_presion { get; set; }
        public string de_rg_presion { get; set; }
        public string cd_metodo_fact { get; set; }
        public string de_metodo_fact { get; set; }
        public string lg_telemedida { get; set; }
        public string cd_tp_gasinera { get; set; }
        public string de_tp_gasinera { get; set; }
        public DateTime fh_desde_reac { get; set; }
        public DateTime fh_hasta_reac { get; set; }
        public string de_marca_back { get; set; }
        public string lg_contrato_simul { get; set; }
        public string cd_id_producto { get; set; }
        public string cd_tp_producto { get; set; }
        public string de_tp_producto { get; set; }
        public string lg_arrastre_penali { get; set; }
        public double nm_med_capacitiva { get; set; }
        public double nm_prec_capacitiva { get; set; }
        public double[] nm_exceso_pot { get; set; }

        public double nm_consumo_medio_5a { get; set; }
        public double nm_consumo_medio { get; set; }


        public FacturasPeajesRedShift()
        {
            nm_med_potencia = new double[10];
            nm_prec_potencia = new double[10];
            nm_med_activa = new double[10];
            nm_prec_activa = new double[10];
            nm_med_reactiva = new double[10];
            nm_prec_reactiva = new double[10];

            cd_concepto = new string[20];
            cd_concepto_sce = new string[20];
            de_concepto = new string[20];
            im_concepto = new double[20];

            de_impuesto = new string[5];
            nm_porcentaje = new double[5];
            nm_base = new double[5];
            nm_importe_impuesto = new double[5];

            consumo_reactiva = new double[6];

            fh_lect_ant_reactiva = new DateTime[6];
            cd_tp_fuente_ant_reactiva = new string[6];
            nm_lect_ant_reactiva = new double[6];
            fh_lect_act_reactiva = new DateTime[6];
            cd_tp_fuente_act_reactiva = new string[6];
            nm_lect_act_reactiva = new double[6];
            im_ajuste_integr_reactiva = new string[6];
            cd_tp_anoma_reactiva = new string[6];

            cd_concepto_reper = new string[12];
            im_concepto_reper = new double[12];

            nm_exceso_pot = new double[6];
        }



    }
}
