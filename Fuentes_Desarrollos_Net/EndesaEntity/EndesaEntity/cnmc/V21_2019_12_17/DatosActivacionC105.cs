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
    public class DatosActivacionC105
    {
        [XmlElement(ElementName = "Fecha", Namespace = "http://localhost/elegibilidad")]
        public string Fecha { get; set; }

        // TABLA PROCESO C2 - ActivacionCambiodeComercializadorConCambios.xsd
        [XmlElement(ElementName = "IndEsencial")] public string IndEsencial { get; set; }

        [XmlElement(ElementName = "FechaUltimoMovimientoIndEsencial")] public string FechaUltimoMovimientoIndEsencial { get; set; }

    }
}
