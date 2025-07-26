using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
   
    [XmlRoot(ElementName = "ActivacionCambiodeComercializadorSinCambios", Namespace = "http://localhost/elegibilidad")]
    [XmlType(Namespace = "http://localhost/elegibilidad")]
    public class ActivacionCambiodeComercializadorSinCambios_C105
    {
        [XmlElement(ElementName = "DatosActivacion", Namespace = "http://localhost/elegibilidad")]
        public DatosActivacionC105 DatosActivacion { get; set; }

        [XmlElement(ElementName = "Contrato", Namespace = "http://localhost/elegibilidad")]
        public Contrato_C105 Contrato { get; set; }
        
        public PuntoDeMedida PuntoDeMedida { get; set; }

        public ActivacionCambiodeComercializadorSinCambios_C105()
        {
            DatosActivacion = new DatosActivacionC105();
        }

    }
}
