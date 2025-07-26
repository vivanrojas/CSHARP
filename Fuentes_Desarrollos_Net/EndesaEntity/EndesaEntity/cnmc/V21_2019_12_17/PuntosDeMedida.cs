using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "PuntosDeMedida")]
    public class PuntosDeMedida
    {
        [XmlElement(ElementName = "PuntoDeMedida")]
        public List<PuntoDeMedida> PuntoDeMedida { get; set; }        
    }
}
