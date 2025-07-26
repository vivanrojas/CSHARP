using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class ClicksPT
    {
        public string nif { get; set; }
        public string cliente { get; set; }
        public string cpe { get; set; }
        public int click { get; set; }        
        public DateTime fecha_operacion { get; set; }
        public string mercado { get; set; }
        public string operacion { get; set; }
        public string producto { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public double bl { get; set; }
        public double fee { get; set; }
        public double volumen { get; set; }
    }
}
