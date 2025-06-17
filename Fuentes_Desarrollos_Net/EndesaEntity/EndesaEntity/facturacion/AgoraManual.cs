using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class AgoraManual
    {
        public string empresa { get; set; }
        public string nif { get; set; }
        public string cliente { get; set; }
        public string cups13 { get; set; }
        public string cups20 { get; set; }
        public string grupo { get; set; }
        public DateTime fecha_vigor_desde { get; set; }
        public DateTime fecha_vigor_hasta { get; set; }
        public bool migrado_sap { get; set; }

        
        
    }
}
