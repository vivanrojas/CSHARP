using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class Dt_peticiones_curva
    {
        public string cups20 { get; set; }
        public string cups13 { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public int dias_facturados_redshift { get; set; }
        public int dias_registrados_redshift { get; set; }

    }
}
