using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "DatosCAU")]
    public class DatosCAUAlta
    { 
        [XmlElement(ElementName = "CAU")] public string CAU { get; set; }
        [XmlElement(ElementName = "TipoAutoconsumo")] public string TipoAutoconsumo { get; set; }
        [XmlElement(ElementName = "TipoSubseccion")] public string TipoSubseccion { get; set; }
        [XmlElement(ElementName = "Colectivo")] public string Colectivo { get; set; }
        public DatosInstGenSolicitud DatosInstGen { get; set; }

        public DatosCAUAlta()
        {
            DatosInstGen = new DatosInstGenSolicitud();
        }
    }

}