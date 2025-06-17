using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class CosPhi
    {
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public double value_from { get; set; }
        public double value_to { get; set; }
        public double value { get; set; }
        public string unit { get; set; }

    }
}
