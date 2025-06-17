using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.gas
{
    public class Cups_Peaje
    {
        public string nif { get; set; }
        public string cliente { get; set; }
        public string cups { get; set; }
        public string peaje { get; set; }
        public string peaje_codigo { get; set; }
        public string distribuidora { get; set; }
        public DateTime fecha_efecto { get; set; }
    }
}
