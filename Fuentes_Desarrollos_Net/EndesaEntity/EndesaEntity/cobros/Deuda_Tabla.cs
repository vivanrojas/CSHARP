using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.cobros
{
    public class Deuda_Tabla
    {

        public string nif { get; set; }
        public string dapersoc { get; set; }
        public DateTime fecha_limite_pago { get; set; }
        public double importe_obligacion { get; set; }
        public DateTime fecha_factura { get; set; }
        public DateTime periodo_factura_desde { get; set; }
        public DateTime periodo_factura_hasta { get; set; }
        public string  cfactura { get; set; }
        public string tipo_factura { get; set; }
        public string cups22 { get; set; }
        public string aajj { get; set; }
        public string agrecobro { get; set; }
        public string pc { get; set; }
        public string fc { get; set; }

        
    }
}
