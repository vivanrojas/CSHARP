using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.gas
{
    public class Informe_GAS_XML
    {
        public string archivo { get; set; }
        public DateTime fecha_importacion { get; set; }
        public string cups_xml { get; set; }
        public string tarifa_xml { get; set; }
        public DateTime fecha_inicio_xml { get; set; }
        public string sistema { get; set; }
        public string cups_sistema { get; set; }
        public string tarifa_sistema { get; set; }
        public DateTime fecha_inicio_sistema { get; set; }
        public bool existe_en_sigame { get; set; }

    }
}
