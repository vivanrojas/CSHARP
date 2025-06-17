using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "DatosActivacion")]
    public class DatosActivacion
    {
        [XmlElement(ElementName = "FechaActivacion")] public string fechaActivacion { get; set; }
        [XmlElement(ElementName = "EnServicio")] public string enServicio { get; set; }
    }
}
