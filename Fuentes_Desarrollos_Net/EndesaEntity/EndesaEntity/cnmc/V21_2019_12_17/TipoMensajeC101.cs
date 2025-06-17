using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{    
    [XmlRoot(ElementName = "MensajeCambiodeComercializadorSinCambios", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    public class TipoMensajeC101
    {
        public Cabecera Cabecera { get; set; }
        public CambioComercializadorSinCambios CambiodeComercializadorSinCambios { get; set; }

        public TipoMensajeC101()
        {
            Cabecera = new Cabecera();
            CambiodeComercializadorSinCambios = new CambioComercializadorSinCambios();
        }
    }
}
