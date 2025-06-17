using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Nombre")]
    public class Nombre
    {
        [XmlElement(ElementName = "NombreDePila")] public string NombreDePila { get; set; }
        [XmlElement(ElementName = "PrimerApellido")] public string PrimerApellido { get; set; }
        [XmlElement(ElementName = "SegundoApellido")] public string SegundoApellido { get; set; }
        [XmlElement(ElementName = "RazonSocial")] public string RazonSocial { get; set; }
    }
}
