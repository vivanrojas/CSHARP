using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    
    public class RegistroDoc
    {
        [XmlElement(ElementName = "TipoDocAportado")] public string TipoDocAportado { get; set; }
        [XmlElement(ElementName = "DireccionUrl")] public string DireccionUrl { get; set; }        
    }
}
