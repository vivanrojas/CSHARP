using EndesaEntity.cnmc.V21_2019_12_17;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.gas.V25_2019_12_17
{
    
    public class CounterList
    {
        [XmlElement(ElementName = "counter")] public List<Counter> counterlist { get; set; }

        public CounterList() 
        {
            counterlist = new List<Counter>();
        }

    }
}
