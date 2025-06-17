using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "IdContrato")]
    public class IdContrato
    {
        [XmlElement(ElementName = "CodContrato")] public string CodContrato { get; set; }
    }
}
