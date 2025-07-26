using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{

    [XmlRoot(ElementName = "MensajeActivacionCambiodeComercializadorSinCambios", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
       public class TipoMensajeC105
    {
        public Cabecera Cabecera { get; set; }

        // [XmlElement(ElementName = "ActivacionCambiodeComercializadorSinCambios", Namespace = "http://localhost/elegibilidad")]
        public ActivacionCambiodeComercializadorSinCambios_C105 ActivacionCambiodeComercializadorSinCambios { get; set; }
        
        public TipoMensajeC105()
        {
            Cabecera = new Cabecera();

        }
    }
}
