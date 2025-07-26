using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class SIGAME_R_20
    {
        public string cfactura { get; set; }
        public string numRefFactura { get; set; }
        public string codConcepto { get; set; }
        public int consumo { get; set; }
        public double importe { get; set; }
        public string descripcion { get; set; }
        public double tipo_impositivo { get; set; }
    }
}
