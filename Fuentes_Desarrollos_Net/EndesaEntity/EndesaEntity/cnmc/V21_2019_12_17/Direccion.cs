using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Direccion")]
    public class Direccion
    {
        [XmlElement(ElementName = "Pais")] public string Pais { get; set; }
        [XmlElement(ElementName = "Provincia")] public string Provincia { get; set; }
        [XmlElement(ElementName = "Municipio")] public string Municipio { get; set; }
        [XmlElement(ElementName = "CodPostal")] public string CodPostal { get; set; }
        
        public Via Via { get; set; }

        public Direccion() 
        { 
            Via = new Via();
        }


    }
}
