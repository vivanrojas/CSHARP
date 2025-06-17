using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Rechazo")]
    public class Rechazo
    {
        [XmlElement(ElementName = "Secuencial")] public string Secuencial { get; set; }
        [XmlElement(ElementName = "CodigoMotivo")] public string CodigoMotivo { get; set; }
        [XmlElement(ElementName = "Comentarios")] public string Comentarios { get; set; }


        public Rechazo()
        {

            
        }
    }
}
