using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Contrato")]
    public class Contrato_B105
    {
        public IdContrato IdContrato { get; set; }

       
    }
}
