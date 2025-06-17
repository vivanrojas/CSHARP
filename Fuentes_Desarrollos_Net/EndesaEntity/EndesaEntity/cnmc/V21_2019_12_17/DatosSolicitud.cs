using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = " DatosSolicitud")]
    public class DatosSolicitud
    {
        [XmlElement(ElementName = "MotivoTraspaso")] public string motivoTraspaso { get; set; }
        [XmlElement(ElementName = "FechaPrevistaAccion")] public string fechaPrevistaAccion { get; set; }
        [XmlElement(ElementName = "SolicitudTension")] public string SolicitudTension { get; set; }
        [XmlElement(ElementName = "CNAE")] public string cnae { get; set; }
        [XmlElement(ElementName = "IndEsencial")] public string indEsencial { get; set; }
        [XmlElement(ElementName = "SuspBajaImpagoEnCurso")] public string suspBajaImpagoEnCurso { get; set; }
        [XmlElement(ElementName = "IndActivacion")] public string IndActivacion { get; set; }
        [XmlElement(ElementName = "Motivo")] public string Motivo { get; set; }


    }
}
