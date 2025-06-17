using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class Consumo_Medio_Mensual
    {
        public string empresa { get; set; }
        public int aniomes { get; set; }
        public string tipo_negocio { get; set; }
        public long consumo { get; set; }
    }
}
