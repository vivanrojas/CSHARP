using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Tunel
    {
        public int id { get; set; }
        public string name { get; set; }
        public DateTime versionStart { get; set; }
        public DateTime versionEnd { get; set; }
        public double volume { get; set; }
        public double lowerBand { get; set; }
        public double higherBand { get; set; }
        public DateTime startDate { get; set; }
        public DateTime endDate { get; set; }
        public string cups20 { get; set; }
        public string cups13 { get; set; }
        public string tarifa { get; set; }
        public double total_energia { get; set; }
        public DateTime fecha_ultima_factura { get; set; }
        public DateTime fecha_primera_factura { get; set; }
        public bool completo { get; set; }
        public string comentario { get; set; }
        public bool es_baja { get; set; }

        public bool tiene_ltp { get; set; }

        
    }
}
