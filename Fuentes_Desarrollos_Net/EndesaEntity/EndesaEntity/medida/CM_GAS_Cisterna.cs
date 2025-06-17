using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class CM_GAS_Cisterna
    {

        public string nombre_cliente { get; set; }
        public int primer_mes_dpte { get; set; }
        public string area_pdte { get; set; }
        public double tam { get; set; }
        public string codigo_slm { get; set; }
        public int mes { get; set; }
        public DateTime fecha_alta { get; set; }

    }
}
