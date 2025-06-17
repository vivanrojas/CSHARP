using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.ree
{
    public class PreciosRegulados
    {
        public string termino { get; set; }
        public string tarifa { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public string unidad { get; set; }
        public double[] periodo_tarifario { get; set; }

        public PreciosRegulados()
        {
            periodo_tarifario = new double[7];
        }
    }
}
