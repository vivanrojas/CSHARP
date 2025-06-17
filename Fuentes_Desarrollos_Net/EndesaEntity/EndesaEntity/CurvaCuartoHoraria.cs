using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class CurvaCuartoHoraria
    {
        public int estacion { get; set; }
        public string cups15 { get; set; }
        public string cups22 { get; set; }
        public string cups20 { get; set; }
        public DateTime fecha { get; set; }
        public int version { get; set; }
        public string estado { get; set; }
        public Int32 totalA { get; set; }
        public Int32 totalR { get; set; }
        public double[] value { get; set; } // potencia Cuartohoraria        
        public double[] a { get; set; }
        public string[] fc { get; set; } // fuente cuartohoraria
        public string[] fa { get; set; } // fuente activa
        public double[] r { get; set; }
        public double[] r1 { get; set; }
        public double[] r2 { get; set; }
        public double[] r3 { get; set; }
        public double[] r4 { get; set; }

        public double AE { get; set; }
        public double AES { get; set; }
        public double R1 { get; set; }
        public double R2 { get; set; }
        public double R3 { get; set; }
        public double R4 { get; set; }
        public double NETEADA { get; set; }

        public DateTime fecha_carga { get; set; }


        public int numPeriodos { get; set; }

        public CurvaCuartoHoraria()
        {
            numPeriodos = 0;
            value = new double[101];
            fc = new string[26];
            a = new double[26];
            fa = new string[26];
            r = new double[26];
            r1 = new double[26];
            r2 = new double[26];
            r3 = new double[26];
            r4 = new double[26];
        }
    }
}
