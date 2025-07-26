using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class Med_cc_horaria
    {
        public int idmp { get; set; }
        public DateTime dia { get; set; }
        public double total_activa { get; set; }
        public double total_reactiva { get; set; }
        public double[] a { get; set; }
        public double[] r { get; set; }
        public string archivo { get; set; }

        public Med_cc_horaria()
        {
            a = new double[25];
            r = new double[25];            
        }

    }
}
