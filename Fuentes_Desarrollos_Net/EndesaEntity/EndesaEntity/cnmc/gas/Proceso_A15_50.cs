using EndesaEntity.cnmc.gas.V25_2019_12_17;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.gas
{
    [XmlRoot(ElementName = "sctdapplication")]
    public class Proceso_A15_50
    {
        public V25_2019_12_17.Heading heading { get; set; }
        public A1550 a1550 { get; set; }

        public Proceso_A15_50()
        {
            heading = new V25_2019_12_17.Heading();
            a1550 = new A1550();
        }
    }
}
