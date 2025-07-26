using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "MensajeRechazo", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    public class TipoMensajeC102_R
    {
        public Cabecera Cabecera { get; set; }
        public Rechazos Rechazos { get; set; }

        public TipoMensajeC102_R()
        {
            Cabecera = new Cabecera();
            Rechazos = new Rechazos();
        }
    }
}
