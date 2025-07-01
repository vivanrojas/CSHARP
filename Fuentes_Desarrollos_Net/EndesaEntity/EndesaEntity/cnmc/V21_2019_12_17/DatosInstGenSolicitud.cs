using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "DatosInstGen")]
    public class DatosInstGenSolicitud
    {
        [XmlElement(ElementName = "CIL")] public string CIL { get; set; }
        //irh
        [XmlElement(ElementName = "TecGenerador")] public string TecGenerador { get; set; }

        [XmlElement(ElementName = "PotInstaladaGen")] public string PotInstaladaGen { get; set; }
        [XmlElement(ElementName = "TipoInstalacion")] public string TipoInstalacion { get; set; }
        [XmlElement(ElementName = "EsquemaMedida")] public string EsquemaMedida { get; set; }
        [XmlElement(ElementName = "SSAA")] public string SSAA { get; set; }
        [XmlElement(ElementName = "UnicoContrato")] public string UnicoContrato { get; set; }
        [XmlElement(ElementName = "RefCatastro")] public string RefCatastro { get; set; }

        // UTM - No implementado porque no es obligatorio
        // TitularRepresentanteGen - No implementado porque no es obligatorio

        public DatosInstGenSolicitud()
        {
            
        }
    }

}