using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class P1
    {
        public string cups22 { get; set; }
        public int tipo_medida { get; set; }
        public string fuente { get; set; }
        public int periodo { get; set; }
        public DateTime fecha_hora { get; set; }        
        public int estacion { get; set; } // 0 invierno 1 verano
        public double ae { get; set; }    // activa entrante
        public int cal_ae { get; set; }   // calidad activa entrante
        public double _as { get; set; }   // activa saliente
        public int cal_as { get; set; }   // calidad activa saliente
        public double[] reactiva {get;set;}
        public int[] cal_reactiva { get; set; }
        public int[] mag_reserva { get; set; }
        public int[] cal_reserva { get; set; }
        public int metodo_obtencion { get; set; }
        public int indicador_firmeza { get; set; } //  0 no firme (R) 1 firme (F)
        public int archivo_id { get; set; }
        public DateTime fecha_archivo { get; set; }


        public P1()
        {
            tipo_medida = 11;
            estacion = 0;
            reactiva = new double[4];
            cal_reactiva = new int[4];
            mag_reserva = new int[2];
            cal_reserva = new int[2];
        }

    }
}
