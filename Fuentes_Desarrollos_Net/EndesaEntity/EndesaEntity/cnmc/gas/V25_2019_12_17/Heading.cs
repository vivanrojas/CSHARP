using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.gas.V25_2019_12_17
{
    [XmlRoot(ElementName = "heading")]
    public class Heading
    {
        [XmlElement(ElementName = "dispatchingcode")] public string dispatchingcode { get; set; }
        [XmlElement(ElementName = "dispatchingcompany")] public string dispatchingcompany { get; set; }
        [XmlElement(ElementName = "destinycompany")] public string destinycompany { get; set; }
        [XmlElement(ElementName = "communicationsdate")] public DateTime communicationsdate { get; set; }
        [XmlElement(ElementName = "communicationshour")] public DateTime communicationshour { get; set; }
        [XmlElement(ElementName = "processcode")] public string processcode { get; set; }
        [XmlElement(ElementName = "messagetype")] public string messagetype { get; set; }

    }
}


