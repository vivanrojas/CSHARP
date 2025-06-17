using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "RegistrosDocumento")]
    public class RegistrosDocumento
    {
        [XmlElement(ElementName = "RegistroDoc")] public List<RegistroDoc> RegistroDoc { get; set; }        

        public RegistrosDocumento()
        {

            RegistroDoc = new List<RegistroDoc>();
        }
    }
}
