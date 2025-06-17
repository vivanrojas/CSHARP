using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Table_atrgas_contratos : FieldsAudit
    {
        public string cnifdnic { get; set; }
        public string dapersoc { get; set; }
        public string distribuidora { get; set; }
        public string cups20 { get; set; }
        public string comentarios_descuadres { get; set; }
        public string comentarios_contratacion { get; set; }
        public string tarifa { get; set; }
        public string tramitacion { get; set; } // Tipo tramitacion para peticiones Distribuidora XML_SCTD, Mail ... etc
    }
}
