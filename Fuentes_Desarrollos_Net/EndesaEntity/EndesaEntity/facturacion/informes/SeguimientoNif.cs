using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion.informes
{
    public class SeguimientoNif
    {

        public string empresa { get; set; }
        public string cif { get; set; }
        public string cliente { get; set; }
        public string cpe { get; set; }
        public string numFactura { get; set; }
        public DateTime fechaFactura { get; set; }
        public DateTime fechaFacturaDesde { get; set; }
        public DateTime fechaFacturaHasta { get; set; }
        public double importeFactura { get; set; }
        public double impuestos { get; set; }
        public string lineaNegocio { get; set; }
        public string tipoFactura { get; set; }
        public string estado { get; set; }

    }
}
