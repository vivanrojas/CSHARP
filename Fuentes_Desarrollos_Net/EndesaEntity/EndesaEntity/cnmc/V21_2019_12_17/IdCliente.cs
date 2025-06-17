using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "IdCliente")]
    public class IdCliente
    {
        [XmlElement(ElementName = "TipoIdentificador")] public string TipoIdentificador { get; set; }
        [XmlElement(ElementName = "Identificador")] public string Identificador { get; set; }
        [XmlElement(ElementName = "TipoPersona")] public string TipoPersona { get; set; }
    }
}
