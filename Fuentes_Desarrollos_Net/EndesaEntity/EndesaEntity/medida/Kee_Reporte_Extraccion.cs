using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class Kee_Reporte_Extraccion
    {
        
        public string cups20 { get; set; }
        public string cups22 { get; set; }
        public string sufijo_cups22 { get; set; }
        public string fuente { get; set; }
        public DateTime fecha { get; set; }
        public string hora { get; set; }
        public int estacion { get; set; }
        public double[] energia { get; set; }
        public int[] cal { get; set; }
        public string metodo_obtencion { get; set; }
        public string indicador_firmeza { get; set; }
        public string archivo { get; set; }

        public Kee_Reporte_Extraccion()
        {
            energia = new double[6];
            cal = new int[6];
        }


    }
}
