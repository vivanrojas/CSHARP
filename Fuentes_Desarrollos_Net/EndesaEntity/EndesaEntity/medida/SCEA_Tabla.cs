using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class SCEA_Tabla
    {
        public string cup13 { get; set; }
        public string cups20 { get; set; }
        public string empresa { get; set; }
        public DateTime fecha_anexion { get; set; }
        public DateTime fecha_prevista_baja { get; set; }
        public string tarifa { get; set; }
        public string estado_contrato { get; set; }
        public string cliente { get; set; }
        public string nif { get; set; }
        public string c3 { get; set; }
        public string end { get; set; }
        public string tipo_distribora { get; set; }
        public string descripcion_distribuidora { get; set; }
        public string tipo_top { get; set; }
        public int tipo_pm_calculado { get; set; }
        public double potencia_maxima_contratada { get; set; }
        public string provincia { get; set; }
        public string comentario_sce { get; set; }
        public DateTime fecha_comentario_sce { get; set; }
        public string info { get; set; }
        public string nombre_gestor { get; set; }
        public string territorial { get; set; }
        public string responsable { get; set; }
        public int pm_pdtes { get; set; }
        public int aaaamm_trabajo { get; set; }
        public int dia_habil { get; set; }
        public int primer_mes_pdte { get; set; }
        public string estado { get; set; }
        public string subestado { get; set; }
        public int num_meses_pdtes { get; set; }
        public int clase { get; set; }
        public string usuario { get; set; }
        public string grupo_resolucion { get; set; }
        public int aaaammlpt { get; set; }
        public string definicion { get; set; }
        public DateTime fechahasta_ltp { get; set; }
        public string anomalia { get; set; }
        public int dificultad { get; set; }
        public string t { get; set; }
        public int activa_ccr { get; set; }
        public int reactiva_ccr { get; set; }
        public string tipo_incompletitud { get; set; }
        public string descripcion_incompletitud { get; set; }
        public int incompletitudes { get; set; }
        public bool existe_FactD { get; set; }
        public DateTime fdesdeccr { get; set; }
        public DateTime fhastaccr { get; set; }
        public int potmaxccr { get; set; }
        public DateTime maxdefh { get; set; }
        public string gestion_propia_atr { get; set; }
        public double afactd { get; set; }
        public double excesospotencia { get; set; }
        public double excesosreactiva { get; set; }
        public string esprimerafactura { get; set; }
        public int aaaamm_ult_fact { get; set; }
        public DateTime fdULTFACT { get; set; }
        public DateTime fhULTFACT { get; set; }
        public double a_ULT_FACT { get; set; }
        public double minA { get; set; }
        public double maxA { get; set; }
        public double medA { get; set; }
        public string peor_fuente { get; set; }
        public string ult_fuente { get; set; }
        public double total_deuda { get; set; }
        public DateTime ult_recep_web { get; set; }
        public string grupo { get; set; }
        public int ranking_tam { get; set; }
        public int ranking_maxfact { get; set; }
        public double tam_por_cups { get; set; }
        public double fact_max { get; set; }
        public int diafactmed { get; set; }
        public int diafacmax { get; set; }
        public string n_pms_cs { get; set; }
        public string tlfcs { get; set; }
        public string tm_ok { get; set; }
        public string consumo_0 { get; set; }
        public string perdidas_ml { get; set; }
        public DateTime maxdefh_Absoluta { get; set; }
        public int aaaamm_maxdefh { get; set; }
        public int aaaamm_maxdefh_absoluta { get; set; }
        public DateTime fecha_baja { get; set; }
        public string proceso_concursal { get; set; }
        public DateTime maxdefh_cierre { get; set; }
        public int aaaamm_maxdefh_cierre { get; set; }

    }
}
