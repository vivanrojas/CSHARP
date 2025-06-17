using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class Med_cc_horaria_tmp
    {
        public string cups22 { get; set; }
        public DateTime dia { get; set; }
        public double total_activa { get; set; }
        public double total_reactiva { get; set; }
        public string[] a { get; set; }
        public string[] r { get; set; }
        public string archivo { get; set; }

        public Med_cc_horaria_tmp()
        {
            a = new string[26];
            r = new string[26];
        }
    }
}
