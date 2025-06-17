using EndesaEntity.cnmc.V21_2019_12_17;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.gas.V25_2019_12_17
{
    [XmlRoot(ElementName = "sctdapplication")]
    public class Proceso_A12_26
    {
        public Heading heading { get; set; }

        public A1226 a1226 { get; set; }

    }
}
