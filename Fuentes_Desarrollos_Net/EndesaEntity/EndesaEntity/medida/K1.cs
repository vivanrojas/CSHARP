using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class K1
    {
        public string cups22 { get; set; }
        public int tipo_medida { get; set; }
        public string fuente { get; set; }
        public string periodo { get; set; }
        public DateTime fecha_hora { get; set; }
        public int estacion { get; set; } // 0 invierno 1 verano
        public string ae { get; set; }    // activa entrante
        public string cal_ae { get; set; }   // calidad activa entrante
        public string _as { get; set; }   // activa saliente
        public string cal_as { get; set; }   // calidad activa saliente
        public string[] reactiva { get; set; }
        public string[] cal_reactiva { get; set; }
        public string[] mag_reserva { get; set; }
        public string[] cal_reserva { get; set; }
        public string metodo_obtencion { get; set; }
        public string indicador_firmeza { get; set; } //  0 no firme (R) 1 firme (F)
        public int archivo_id { get; set; }
        public string archivo { get; set; }
        public DateTime fecha_archivo { get; set; }


        public K1()
        {
            tipo_medida = 11;
            estacion = 0;
            reactiva = new string[4];
            cal_reactiva = new string[4];
            mag_reserva = new string[2];
            cal_reserva = new string[2];
        }
    }
}
