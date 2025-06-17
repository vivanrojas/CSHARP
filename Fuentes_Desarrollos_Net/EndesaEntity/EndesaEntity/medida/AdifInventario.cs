using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class AdifInventario : Audit
    {
        public int num { get; set; }
        public int id_cups { get; set; }
        public string cups13 { get; set; }
        public string cups20 { get; set; }
        public int lote { get; set; }
        public DateTime vigencia_desde { get; set; }
        public DateTime vigencia_hasta { get; set; }
        public string zona { get; set; }
        public string codigo { get; set; }
        public string nombre_punto_suministro { get; set; }
        public string distribuidora { get; set; }
        public string comentarios { get; set; }
        public string tarifa { get; set; }
        public int tension { get; set; }
        public double[] p { get; set; }
        public string estado_curva { get; set; }
        public int dias { get; set; }
        public DateTime ffactdes { get; set; }
        public DateTime ffacthas { get; set; }
        public string fuente { get; set; }
        public double resumen_sce_a { get; set; }
        public double adif_a { get; set; }
        public double sce_a { get; set; }
        public double ree_a { get; set; }
        public double dif_sce_adif_a { get; set; }
        public double resumen_sce_r { get; set; }
        public double adif_r { get; set; }
        public double sce_r { get; set; }
        public double ree_r { get; set; }
        public double dif_sce_adif_r { get; set; }
        public string estado_contrato { get; set; }
        public string resultado { get; set; }
        public string enviado_facturar { get; set; }
        public bool medida_en_baja { get; set; }
        public bool devolucion_de_energia { get; set; }
        public bool cierres_energia { get; set; }
        public double activa_entrante { get; set; }
        public double activa_saliente { get; set; }    
        public double neteada { get; set; }
        public double valor_kvas { get; set; }
        public  int multipunto_num_principales { get; set; }
        public string provincia { get;set; }
        public string comunidad_autonoma { get; set; }
        public double perdidas { get; set; }
        public string grupo { get; set; }
        public string sitema_traccion { get; set; }
        
        public bool tiene_perdidas { get; set; }

        public List<CurvaCuartoHorariaTabla> curvaHorariaADIF {get;set;}
        public List<CurvaResumenTabla> curvaResumen { get; set; }

        public AdifInventario()
        {
            p = new double[6];
            curvaResumen = new List<CurvaResumenTabla>();
            curvaHorariaADIF = new List<CurvaCuartoHorariaTabla>();
        }
    }
}
