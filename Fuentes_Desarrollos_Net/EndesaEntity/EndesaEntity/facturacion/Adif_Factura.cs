using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace EndesaEntity.facturacion
{
    public class Adif_Factura
    {
        public string clave { get; set; }
        public String CUPSREE { get; set; }
        public String CUPS13 { get; set; }
        public int LOTE { get; set; }
        public String COMENTARIO { get; set; }
        public String EMPRESA { get; set; }
        public string CREFEREN { get; set; }
        public int SECFACTU { get; set; }
        public DateTime FFACTURA { get; set; }
        public DateTime FFACTDES { get; set; }
        public DateTime FFACTHAS { get; set; }
        public Int64 CONSUMO_ADIF { get; set; }
        public Int64 CONSUMO_SCE { get; set; }
        public double DIF_CONSUMO { get; set; }
        public String PRODUCTO { get; set; }
        public String PRECIO { get; set; }
        public Double VE { get; set; }
        public Double KE { get; set; }
        public Double IMPORTE_CON_IE { get; set; }
        public Double BASE_ISE { get; set; }
        public String PCT_IVA { get; set; }
        public Double IVA { get; set; }
        public Double TOTAL_ADIF { get; set; }
        public Double TOTAL_SCE { get; set; }
        public Double DIF_TOTAL { get; set; }
        public string FACTURADO { get; set; }
        public string cfactura { get; set; }
        public double cnpr_adif { get; set; }
        public double cnpr_endesa {get; set; }
        public double cpre_adif { get; set; }
        public double cpre_endesa { get; set; }
        public string testfact { get; set; }
        public string de_tfactura { get; set; }
        public double importe_conceptos { get; set; }
        public double ifactura { get; set; }
        public double ise { get; set; }
        public bool medida_en_baja { get; set; }
        public bool devolucion_de_energia { get; set; }
        public bool cierres_energia { get; set; }
        public bool existe_factura_adif { get; set; }
        public bool existe_factura_sce { get; set; }



    }
}
