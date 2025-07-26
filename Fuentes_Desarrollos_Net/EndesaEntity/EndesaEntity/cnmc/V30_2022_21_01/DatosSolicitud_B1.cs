using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V30_2022_21_01
{
    [XmlRoot(ElementName = " DatosSolicitud")]
    public class DatosSolicitud_B1
    {
        [XmlElement(ElementName = "IndActivacion")] public string IndActivacion { get; set; }
        [XmlElement(ElementName = "FechaPrevistaAccion")] public string fechaPrevistaAccion { get; set; }
        [XmlElement(ElementName = "Motivo")] public string Motivo { get; set; }        
    }
}
