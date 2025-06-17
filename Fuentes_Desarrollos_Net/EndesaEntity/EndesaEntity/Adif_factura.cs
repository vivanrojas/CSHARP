using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Adif_factura
    {
        public String CUPSREE { get; set; }

        public String CUPS13 { get; set; }
        public int LOTE { get; set; }
        public String COMENTARIO { get; set; }
        public String EMPRESA { get; set; }
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
    }
}
