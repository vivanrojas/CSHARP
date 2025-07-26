using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.factoring
{
    public class ListaNegra : Audit
    {
        public string nif { get; set; }
        public string cliente { get; set; }
    }

    public class ListaNegra_CUPS: ListaNegra
    {
        public string cups20 { get; set; }
        public DateTime fecha_inicio { get; set; }
        public DateTime fecha_fin { get; set; }
        public string motivo { get; set; }
    }
}
