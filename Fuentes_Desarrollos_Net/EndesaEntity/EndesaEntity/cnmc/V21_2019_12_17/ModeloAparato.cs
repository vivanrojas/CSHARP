using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "ModeloAparato")]
    public class ModeloAparato
    {
        [XmlElement(ElementName = "TipoAparato")] public string TipoAparato { get; set; }
        [XmlElement(ElementName = "MarcaAparato")] public string MarcaAparato { get; set; }
        [XmlElement(ElementName = "ModeloMarca")] public string ModeloMarca { get; set; }
    }
}
