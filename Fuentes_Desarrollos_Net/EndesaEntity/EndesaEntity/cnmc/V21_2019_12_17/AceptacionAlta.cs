using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "AceptacionAlta")]
    public class AceptacionAlta
    {
        public DatosAceptacion DatosAceptacion { get; set; }
        public ContratoAlta Contrato { get; set; }

        public AceptacionAlta()
        {
            DatosAceptacion = new DatosAceptacion();
            Contrato = new ContratoAlta();
        }
    }
}
