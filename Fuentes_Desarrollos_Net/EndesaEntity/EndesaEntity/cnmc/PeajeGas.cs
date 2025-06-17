using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.cnmc
{
    public class PeajeGas
    {
        public string tarifa { get; set; }
        public string codigo { get; set; }
        public int grupo_presion { get; set; }
        public double consumo_anual_desde { get; set; }
        public double consumo_anual_hasta { get; set; }
        public bool gas_xxi { get; set; }
    }
}
