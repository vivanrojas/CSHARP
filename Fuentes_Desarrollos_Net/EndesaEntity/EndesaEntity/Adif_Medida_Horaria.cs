using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Adif_Medida_Horaria
    {
        public int id { get; set; }
        public string cups13 { get; set; }
        public string cup20 { get; set; }
        public DateTime fecha { get; set; }
        public string tipoEnergia { get; set; }
        public string unidad { get; set; }
        public Int32 total { get; set; }
        public Int32[] value { get; set; }
        public string[] c { get; set; }
        public DateTime fechaCarga { get; set; }
        public string fichero { get; set; }
        public int numPeriodos { get; set; }

        public Adif_Medida_Horaria()
        {
            value = new Int32[26];
            c = new string[26];

        }
    }
}
