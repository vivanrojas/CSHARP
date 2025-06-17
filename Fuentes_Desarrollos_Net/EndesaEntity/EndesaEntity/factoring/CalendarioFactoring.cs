using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.factoring
{
    public class CalendarioFactoring
    {
        public string factoring { get; set; }
        public int bloque { get; set; }
        public DateTime facturas_desde { get; set; }
        public DateTime facturas_hasta { get; set; }
        public DateTime consumos_desde { get; set; }
        public DateTime consumos_hasta { get; set; }
        public DateTime fecha_ejecucion_desde { get; set; }
        public DateTime fecha_ejecucion_hasta { get; set; }
        public bool ejecutado { get; set; }
        public double importe_min_factura { get; set; }
        public double importe_min_factura_agrupada { get; set; }
    }
}
