using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class PuntosSofisticados : Audit
    {
        public string cups13 { get; set; }
        public string cups20 { get; set; }
        public string dapersoc { get; set; }
        public string grupo { get; set; }
        public DateTime fd { get; set; }
        public DateTime fh { get; set; }
        public string precios { get; set; }
        public bool facturas_a_cuenta { get; set; }

    }
}
