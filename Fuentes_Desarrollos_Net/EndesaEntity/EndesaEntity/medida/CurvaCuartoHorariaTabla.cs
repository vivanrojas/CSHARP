using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class CurvaCuartoHorariaTabla
    {
        public string cups15 { get; set; }
        public string cups22 { get; set; }
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
        public int numPeriodos { get; set; }

        public CurvaCuartoHorariaTabla()
        {
            numPeriodos = 0;
            value = new double[100];
            fc = new string[25];
            a = new double[25];
            fa = new string[25];
            r = new double[25];
        }
    }

}
