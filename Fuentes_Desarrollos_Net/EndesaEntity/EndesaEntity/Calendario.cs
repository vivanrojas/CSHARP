using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Calendario
    {
        public string tarifa { get; set; }
        public string territorio { get; set; }
        public string estacion { get; set; }
        public bool laborable { get; set; }
        public DateTime fd { get; set; }
        public DateTime fh { get; set; }
        public int[] periodos { get; set; }

        public Calendario()
        {
            periodos = new int[25];
        }
    }
}
