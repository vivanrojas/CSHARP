using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "AceptacionModificacionDeATR", Namespace = "http://localhost/elegibilidad")]

    [XmlType(Namespace = "http://localhost/elegibilidad")]
    public class AceptacionModificionDeATR
    {
        public DatosAceptacion DatosAceptacion { get; set; }
        public Contrato_C102_A Contrato { get; set; }

        public AceptacionModificionDeATR()
        {
           DatosAceptacion = new DatosAceptacion();
           Contrato = new Contrato_C102_A();
        }
    }
}
