using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Telefono")]
    public class Telefono
    {
        [XmlElement(ElementName = "PrefijoPais")] public string PrefijoPais { get; set; }
        [XmlElement(ElementName = "Numero")] public string Numero { get; set; }
        [XmlElement(ElementName = "CorreoElectronico")] public string CorreoElectronico { get; set; }
    }
}
