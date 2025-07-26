using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Rechazos")]
    public class Rechazos
    {
        [XmlElement(ElementName = "FechaRechazo")] public string FechaRechazo { get; set; }
        [XmlElement(ElementName = "Rechazo")] public List<Rechazo> Rechazo { get; set; }
        public RegistrosDocumento RegistrosDocumento { get; set; }

        public Rechazos()
        {
            Rechazo = new List<Rechazo>();
        }
    }
}
