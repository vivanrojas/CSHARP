using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class SIGAME_Informe_IH
    {
        public string cliente { get; set; }
        public string nif { get; set; }
        public int id_ps { get; set; }
        public string cups { get; set; }
        public string direccion_punto_suministro { get; set; }
        public string provincia { get; set; }
        public string cae { get; set; }
        public bool ihtc { get; set; }
        public bool ihtd {get;set;}
        public bool ihti { get; set; }
        public bool exento { get; set; }

    }
}
