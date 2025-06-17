using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.gas
{
    public class Addenda
    {
        public int id_ps { get; set; }
        public  string cups { get; set; }
        public string tarifa_peaje { get; set; }
        public string duracion_peaje { get; set; }
        public double qd { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }

    }
}
