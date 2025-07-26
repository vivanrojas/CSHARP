using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = " DatosAceptacion", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    public class DatosAceptacionB102
    {
        [XmlElement(ElementName = "FechaAceptacion")] public string FechaAceptacion { get; set; }
        [XmlElement(ElementName = "ActuacionCampo")] public string ActuacionCampo { get; set; }
        //irh
        [XmlElement(ElementName = "FechaUltimaLecturaFirme")] public string FechaUltimaLecturaFirme { get; set; }
    
        [XmlElement(ElementName = "TipoActivacionPrevista")] public string TipoActivacionPrevista { get; set; }
        [XmlElement(ElementName = "FechaActivacionPrevista")] public string FechaActivacionPrevista { get; set; }
        
    }
}
