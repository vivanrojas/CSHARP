using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    
    [XmlRoot(ElementName = "DatosActivacion", Namespace = "http://localhost/elegibilidad")]
    [XmlType(Namespace = "http://localhost/elegibilidad")]
    public class DatosActivacion
    {
        [XmlElement(ElementName = "FechaActivacion")] public string fechaActivacion { get; set; }
        [XmlElement(ElementName = "EnServicio")] public string enServicio { get; set; }

        //irh
        [XmlElement(ElementName = "Fecha")] public string fecha { get; set; }
        [XmlElement(ElementName = "IndEsencial")] public string IndEsencial { get; set; }
        [XmlElement(ElementName = "FechaUltimoMovimientoIndEsencial")] public string FechaUltimoMovimientoIndEsencial { get; set; }

    }
}
