using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class PendienteSAPKEE
    {

        public string cups { get; set; }
        public string mes { get; set; }

        public DateTime ult_fecha_desde_facturada { get; set; }
        public DateTime ult_fecha_hasta_facturada { get; set; }

        public DateTime fecha_factura { get; set; }
        

    }
}

