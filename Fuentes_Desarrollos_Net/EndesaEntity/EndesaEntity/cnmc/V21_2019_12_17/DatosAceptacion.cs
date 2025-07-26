using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = " DatosAceptacion")]
    public class DatosAceptacion
    {
        [XmlElement(ElementName = "FechaAceptacion")] public string fechaAceptacion { get; set; }
        [XmlElement(ElementName = "ActuacionCampo")] public string ActuacionCampo { get; set; }
        //irh
        [XmlElement(ElementName = "FechaUltimaLecturaFirme")] public string FechaUltimaLecturaFirme { get; set; }
    }
}
