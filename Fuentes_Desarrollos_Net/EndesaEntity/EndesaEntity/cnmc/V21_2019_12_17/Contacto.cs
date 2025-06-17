using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Contacto")]
    public class Contacto
    {
        [XmlElement(ElementName = "PersonaDeContacto")] public string PersonaDeContacto { get; set; }
        public Telefono Telefono { get; set; }
        public Contacto()
        {
            Telefono = new Telefono();
        }
    }
}
