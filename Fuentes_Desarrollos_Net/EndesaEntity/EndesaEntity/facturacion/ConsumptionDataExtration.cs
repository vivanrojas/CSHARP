using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class ConsumptionDataExtration
    {
        public int anio { get; set; }
        public int mes { get; set; }
        public string mercado { get; set; }
        public string linea { get; set; }
        public string segmento { get; set; }
        public Int32 consumo { get; set; }
    }
}
