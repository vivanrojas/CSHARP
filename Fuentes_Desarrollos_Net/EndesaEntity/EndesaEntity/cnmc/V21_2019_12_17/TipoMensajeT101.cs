using EndesaEntity.facturacion.cuadroDeMando;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "MensajeSolicitudTraspasoCOR", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    
    public class TipoMensajeT101
    {
        public Cabecera Cabecera { get; set; }
        public SolicitudTraspasoCOR SolicitudTraspasoCOR { get; set; }
        

        public TipoMensajeT101()
        {
            Cabecera = new Cabecera();
            SolicitudTraspasoCOR = new SolicitudTraspasoCOR();
            
        }

    }
}
