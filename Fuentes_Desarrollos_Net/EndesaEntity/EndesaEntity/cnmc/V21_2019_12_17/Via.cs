using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Via")]
    public class Via
    {
        [XmlElement(ElementName = "TipoVia")] public string TipoVia { get; set; }
        [XmlElement(ElementName = "Calle")] public string Calle { get; set; }
        [XmlElement(ElementName = "NumeroFinca")] public string NumeroFinca { get; set; }
        [XmlElement(ElementName = "DuplicadorFinca")] public string DuplicadorFinca { get; set; }
        [XmlElement(ElementName = "Escalera")] public string Escalera { get; set; }
        [XmlElement(ElementName = "Piso")] public string Piso { get; set; }
        [XmlElement(ElementName = "Puerta")] public string Puerta { get; set; }
        [XmlElement(ElementName = "TipoAclaradorFinca")] public string TipoAclaradorFinca { get; set; }
        [XmlElement(ElementName = "AclaradorFinca")] public string AclaradorFinca { get; set; }
    }
}
