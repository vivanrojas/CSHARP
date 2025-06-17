using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EndesaEntity.cnmc.gas.V25_2019_12_17
{
    public class Informe_Reubicaciones
    {
        public string cups { get; set; }
        public string tipo { get; set; }
        public string tarifa_sigame { get; set; }
        public string tarifa_reubicacion { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }

        public DateTime fecha_reubicacion { get; set; }



    }
}
