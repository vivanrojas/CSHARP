using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{

    [XmlRoot(ElementName = " DatosSolicitud")]
    public class DatosSolicitud_A3
    {
        [XmlElement(ElementName = "CNAE")] public string cnae { get; set; }
        // irh 
        [XmlElement(ElementName = "IndEsencial")] public string IndEsencial { get; set; }
        
        [XmlElement(ElementName = "IndActivacion")] public string IndActivacion { get; set; }
        [XmlElement(ElementName = "FechaPrevistaAccion")] public string fechaPrevistaAccion { get; set; }
        [XmlElement(ElementName = "SolicitudTension")] public string SolicitudTension { get; set; }         
        [XmlElement(ElementName = "TensionSolicitada")] public string TensionSolicitada  { get; set; }        
        
    }
}
