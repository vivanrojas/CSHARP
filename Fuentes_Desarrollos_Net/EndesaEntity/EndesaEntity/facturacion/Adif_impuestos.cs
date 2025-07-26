using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class Adif_impuestos
    {
        public string tipoImpuesto { get; set; }
        public string descripcion { get; set; }
        public DateTime fechaDesde { get; set; }
        public DateTime fechaHasta { get; set; }
        public double valor { get; set; }
        public string unidad { get; set; }
    }
}
