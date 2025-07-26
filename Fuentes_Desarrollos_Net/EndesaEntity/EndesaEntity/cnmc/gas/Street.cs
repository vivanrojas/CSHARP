using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.gas
{
    [XmlRoot(ElementName = "street")]
    public class Street
    {
        [XmlElement(ElementName = "streettype")] public string streettype { get; set; }
        [XmlElement(ElementName = "street")] public string street { get; set; }
        [XmlElement(ElementName = "streetnumber")] public string streetnumber { get; set; }        
        [XmlElement(ElementName = "floor")] public string floor { get; set; }
        [XmlElement(ElementName = "door")] public string door { get; set; }
    }
}

