using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.cnmc.gas.V25_2019_12_17
{
    public class Cargas
    {
        public DateTime fecha_carga { get; set; }
        public string num_archivos { get; set; }
        public string fecha_minima_solicitud { get; set; }
        public string fecha_maxima_solicitud { get; set; }
    }
}
