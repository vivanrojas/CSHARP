using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.xxi
{
    public class Informe_Facturacion
    {
        public string codigo_solicitud { get; set; }
        public string cups { get; set; }
        public string nif { get; set; }
        public string cliente { get; set; }
        public DateTime fecha_alta { get; set; }
        public DateTime fecha_baja { get; set; }
        public DateTime alta_ee { get; set; }
        public int dif { get; set; }
        public double p1 { get; set; }
        public double p2 { get; set; }
        public double p3 { get; set; }
        public double p4 { get; set; }
        public double p5 { get; set; }
        public double p6 { get; set; }

        public string tarifa { get; set; }
        public double tension { get; set; }
        public DateTime fecha_envio_mail { get; set; }

      
    }
}
