using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class DatosPeaje
    {

        public string cups20 { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public int version { get; set; }
        public double[] activa { get; set; }
        public double total_energia_activa { get; set; }
        public double total_energia_reactiva {get;set;}        
        public double[] reactiva { get; set; }
        public double[] potmax { get; set; }
        public double importe_termino_potencia { get; set; }
        public double importe_excesos_potencia { get; set; }
        public double importe_excesos_reactiva { get; set; }

        public DatosPeaje()
        {
            activa = new double[7];
            reactiva = new double[7];
            potmax = new double[7];
        }

    }
}
