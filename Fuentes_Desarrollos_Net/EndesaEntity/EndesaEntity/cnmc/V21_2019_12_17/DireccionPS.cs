using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "DireccionPS")]
    public class DireccionPS
    {
        [XmlElement(ElementName = "Pais")] public string Pais { get; set; }
        [XmlElement(ElementName = "Provincia")] public string Provincia { get; set; }
        [XmlElement(ElementName = "Municipio")] public string Municipio { get; set; }
        [XmlElement(ElementName = "TipoVia")] public string TipoVia { get; set; }
        [XmlElement(ElementName = "CodPostal")] public string CodPostal { get; set; }
        [XmlElement(ElementName = "Calle")] public string Calle { get; set; }
        [XmlElement(ElementName = "NumeroFinca")] public string NumeroFinca { get; set; }

    }
}
