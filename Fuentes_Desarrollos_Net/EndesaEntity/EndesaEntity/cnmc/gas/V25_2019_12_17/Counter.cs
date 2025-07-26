using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.gas.V25_2019_12_17
{

    [XmlRoot(ElementName = "counter")]
    public class Counter 
    {
        [XmlElement(ElementName = "countermodel")] public string countermodel { get; set; }
        [XmlElement(ElementName = "countertype")] public string countertype { get; set; }
        [XmlElement(ElementName = "counternumber")] public string counternumber { get; set; }
        [XmlElement(ElementName = "counterproperty")] public string counterproperty { get; set; }
        [XmlElement(ElementName = "reallecture")] public double reallecture { get; set; }
        [XmlElement(ElementName = "counterpressure")] public double counterpressure { get; set; }
    }
}
