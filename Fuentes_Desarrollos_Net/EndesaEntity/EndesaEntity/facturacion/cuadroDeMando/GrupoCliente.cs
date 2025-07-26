using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion.cuadroDeMando
{
    public class GrupoCliente: Audit
    {
        public string grupo { get; set; }
        public string cliente { get; set; }
        public string cups20 { get; set; }
        public DateTime fecha_desde { get; set; }

        public DateTime fecha_hasta { get; set; }

    }
}
