using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    
    public class Potencia   
    {

        [XmlAttribute(AttributeName = "Periodo")]         
        public string periodo { get; set; }
        [XmlText] public string potencia { get; set; }

       

    }
}
