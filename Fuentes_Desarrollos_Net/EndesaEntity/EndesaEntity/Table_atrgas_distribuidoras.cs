using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Table_atrgas_distribuidoras : FieldsAudit    
    {
        public string codigo { get; set; }
        public string nombre { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public string email { get; set; }
        public string tramitacion { get; set; }
        public string codigo_xml { get; set; }


    }
}
