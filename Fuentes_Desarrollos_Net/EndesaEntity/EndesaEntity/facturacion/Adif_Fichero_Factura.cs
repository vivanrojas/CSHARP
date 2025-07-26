using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class Adif_Fichero_Factura
    {
        public Int64 id_fichero { get; set; }
        public Int64 id_cups_lote { get; set; }
        public String cupsree { get; set; }
        public Int32 lote { get; set; }
    }
}
