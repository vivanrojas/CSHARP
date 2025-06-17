using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class AdifInventario
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
        public int[] p { get; set; }
        public string estado_curva { get; set; }
        public int dias { get; set; }
        public DateTime ffactdes { get; set; }
        public DateTime ffacthas { get; set; }
        public string fuente { get; set; }
        public Int32 resumen_sce_a { get; set; }
        public Int32 adif_a { get; set; }
        public Int32 sce_a { get; set; }
        public Int32 ree_a { get; set; }
        public Int32 dif_sce_adif_a { get; set; }
        public Int32 resumen_sce_r { get; set; }
        public Int32 adif_r { get; set; }
        public Int32 sce_r { get; set; }
        public Int32 ree_r { get; set; }
        public Int32 dif_sce_adif_r { get; set; }
        public string estado_contrato { get; set; }
        public string resultado { get; set; }
        public string enviado_facturar { get; set; }
        public bool medida_en_baja { get; set; }
        public bool devolucion_de_energia { get; set; }
        public bool cierres_energia { get; set; }



        public AdifInventario()
        {
            p = new int[5];
        }
    }
}
