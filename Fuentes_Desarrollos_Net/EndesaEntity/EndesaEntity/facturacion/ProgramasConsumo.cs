using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class ProgramasConsumo
    {
        public string cups20 { get; set; }
        public DateTime fecha { get; set; }
        public int mercado { get; set; }
        public string unidad { get; set; }
        public double[] value { get; set; }

        public ProgramasConsumo()
        {
            value = new double[26];
        }



    }
}
