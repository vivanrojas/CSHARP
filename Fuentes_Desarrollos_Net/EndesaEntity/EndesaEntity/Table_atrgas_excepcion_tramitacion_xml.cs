using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Table_atrgas_excepcion_tramitacion_xml : FieldsAudit
    {
        public Int32 id { get; set; }
        public string nombre { get; set; }
        public string tramitacion { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        
        // El campo estado está en la tabla atrgas_excepcion_tramitacion_xml, será actualizado en código cuando proceda,
        // usado exclusivamente para propósitos de visualización en formulario
        // La excepción podrá estar: Programada, En ejecución, Finalizada o Cancelada
        public string estado { get; set; }
        
    }
}
