using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class PreciosEnergia
    {
        public string cups20 { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public double[] precios_periodo { get; set; }
        public double descuento_te { get; set; }

        public PreciosEnergia()
        {
            precios_periodo = new double[7];
        }
    }
}
