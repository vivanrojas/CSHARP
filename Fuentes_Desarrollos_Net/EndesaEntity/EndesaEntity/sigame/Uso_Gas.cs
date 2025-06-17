using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.sigame
{
    public class Uso_Gas
    {
        public int id_ps { get; set; }
        public int id_pmedida { get; set; }
        public DateTime fd { get; set; }
        public DateTime fh { get; set; }
        public string uso_gas { get; set; }
        public double porcentaje_uso_gas { get; set; }
    }
}
