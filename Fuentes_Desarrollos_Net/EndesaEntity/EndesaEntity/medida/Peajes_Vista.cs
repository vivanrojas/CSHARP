using EndesaEntity.extrasistemas;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class Peajes_Vista
    {

        public string cups13 { get; set; }
        public string cups20 { get; set; }
        public string cod_factura { get; set; }
        public string estado_factura { get; set; }
        public string tipo_factura { get; set; }
        public DateTime fecha_factura { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public string ind_cuadre_periodo_c { get; set; }
        public double ffm_ae_total_c { get; set; }
        public double ffm_r1_total_c { get; set; }
        public double ffm_pmax_total_c { get; set; }
        public double fff_ae_total_c { get; set; }
        public double fff_r1_total_c { get; set; }
        public double fff_pfac_total_c { get; set; }
        public double fff_imp_exc_pot_total_c { get; set; }
        public double fff_imp_exc_pot { get; set; }
        public double fff_imp_exc_r1 { get; set; }
        public double[] ffm_ae { get; set; }
        public double[] ffm_r1 { get; set; }
        public double[] ffm_pmax { get; set; }
        public double[] fff_ae { get; set; }
        public double[] fff_r1 { get; set; }
        public double fff_r4_6 { get; set; }
        public double[] fff_pfac { get; set; }
        public double[] fff_imp_exc_pot1 { get; set; }
        public string cups20_metra { get; set; }
        public string procedencia { get; set; }
        public string ind_perdidas { get; set; }
        public double porcentaje_perdidas { get; set; }
        public double pot_trafo_perdidas_va { get; set; }
        public string tipo_pm { get; set; }
        public string ind_autoconsumo { get; set; }
        public string ind_telemedida { get; set; }
        public string cod_factura_sustituida { get; set; }
        public string cod_factura_rectificada { get; set; }
        public string codigo_estado_factura { get; set; }
        public string tarifa { get; set; }
        public string tipo_consumo { get; set; }
        public string tipo_telegestion { get; set; }
        public string contrato_ext_ps { get; set; }
        public string contrato_ps { get; set; }
        public int sec_factura { get; set; }
        public string distribuidora { get; set; }
        public double consumo_tot_act { get; set; }
        public double consumo_tot_react { get; set; }
        public DateTime fecha_desde_ae { get; set; }
        public DateTime fecha_hasta_ae { get; set; }
        public DateTime fecha_desde_r1 { get; set; }
        public DateTime fecha_hasta_r1 { get; set; }
        public DateTime fecha_desde_pfac { get; set; }
        public DateTime fecha_hasta_pfac { get; set; }
        public DateTime fecha_desde_curva { get; set; }
        public DateTime fecha_hasta_curva { get; set; }
        public string cups_ext { get; set; }
        public int cod_carga_ods { get; set; }
        public DateTime fecha_act_ods { get; set; }
        public DateTime fecha_act_dmco { get; set; }
        public DateTime fecha_recepcion { get; set; }
        public int cod_carga { get; set; }

        public Peajes_Vista()
        {
            ffm_ae = new double[7];
            ffm_r1 = new double[7];
            ffm_pmax = new double[7];
            fff_ae = new double[7];
            fff_r1 = new double[7];
            fff_pfac = new double[7];
            fff_imp_exc_pot1 = new double[7];
        }    




    }
}
