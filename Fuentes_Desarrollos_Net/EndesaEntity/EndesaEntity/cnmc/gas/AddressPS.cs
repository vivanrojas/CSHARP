using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.gas
{

    [XmlRoot(ElementName = "addressPS")]
    public class AddressPS
    {
        [XmlElement(ElementName = "province")] public string province { get; set; }
        [XmlElement(ElementName = "city")] public string city { get; set; }
        [XmlElement(ElementName = "zipcode")] public string zipcode { get; set; }
        public Street street { get; set; }
        public AddressPS()
        {
            street = new Street();
        }
    }
}
