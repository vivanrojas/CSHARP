using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    
    [XmlRoot(ElementName = "DatosActivacionBaja", Namespace = "http://localhost/elegibilidad")]
    [XmlType(Namespace = "http://localhost/elegibilidad")]
    public class DatosActivacionBaja
    {
        [XmlElement(ElementName = "FechaActivacion")] public string FechaActivacion { get; set; }
        

    }
}
