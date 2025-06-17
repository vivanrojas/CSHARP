using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.factoring
{
    public class Estimacion
    {
        public int factoring { get; set; }
        public string tipo { get; set; }
        public string ln { get; set; }
        public int empresa_titular { get; set; }
        public string nif { get; set; }
        public string cliente { get; set; }
        public string ccounips { get; set; }
        public string cupsree { get; set; }
        public string referencia { get; set; }
        public int sec { get; set; }
        public int control { get; set; }
        public double estimacion_importe { get; set; }
        public double estimacion_base { get; set; }
        public double estimacion_impuestos { get; set; }
        public DateTime diaF { get; set; }
        public DateTime diaV { get; set; }
        public double tam { get; set; }
        public double[] ifactura { get; set; }
        public double[] impuestos { get; set; }
        public int[] f { get; set; }
        public int[] dvto { get; set; }

        public Estimacion()
        {
            ifactura = new double[4];
            impuestos = new double[4];
            f = new int[4];
            dvto = new int[4];
        }
    }
}
