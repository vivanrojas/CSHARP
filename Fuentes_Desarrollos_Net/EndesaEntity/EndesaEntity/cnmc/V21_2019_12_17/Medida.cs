using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{

    
    public class Medida
    {
        [XmlElement(ElementName = "TipoDHEdM")] public string TipoDHEdM { get; set; }
        [XmlElement(ElementName = "Periodo")] public string Periodo { get; set; }
        [XmlElement(ElementName = "MagnitudMedida")] public string MagnitudMedida { get; set; }
        [XmlElement(ElementName = "Procedencia")] public string Procedencia { get; set; }
        [XmlElement(ElementName = "UltimaLecturaFirme")] public string UltimaLecturaFirme { get; set; }
        [XmlElement(ElementName = "FechaLecturaFirme")] public string FechaLecturaFirme { get; set; }
        [XmlElement(ElementName = "Anomalia")] public string Anomalia { get; set; }
        [XmlElement(ElementName = "Comentarios")] public string Comentarios { get; set; }
        
    }
}


