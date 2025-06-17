using EndesaEntity.cnmc.V21_2019_12_17;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V30_2022_21_01
{
    [XmlRoot(ElementName = "MensajeCambiodeComercializadorConCambios", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    //[XmlRoot(ElementName = "CambiodeComercializadorConCambios", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    public class TipoMensajeC201
    {
        public Cabecera Cabecera { get; set; }

        [XmlElement(ElementName = "CambiodeComercializadorConCambios", Namespace = "http://localhost/elegibilidad")]
                                  
        public CambioComercializadorConCambios CambioComercializadorConCambios  { get; set; }
        public TipoMensajeC201() 
        {
           Cabecera = new Cabecera();
           CambioComercializadorConCambios = new CambioComercializadorConCambios();
        
        }

    }
}
