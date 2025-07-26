using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion.cuadroDeMando
{
    public class NRI
    {
        public DateTime fecha_ultimo_estado { get; set; }
        public string estado { get; set; }

        public int codigo_nri { get; set; }
        public string cliente { get; set; }
        public string nif { get; set; }
        public string plazo { get; set; }

        public string motivo_alta { get; set; }
        public string submotivo_alta { get; set; }
        public string linea_negocio { get; set; }

    }
}
