using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class SIGAME_TablaBaseDetalle
    {
        public int consumo_parcial { get; set; }
        public string cod_concepto { get; set; }
        public string tipo_impositivo_ih { get; set; }
        public double importe_ih { get; set; }
        public double importe_ih_parcial { get; set; }
        public string descripcion { get; set; }
    }
}
