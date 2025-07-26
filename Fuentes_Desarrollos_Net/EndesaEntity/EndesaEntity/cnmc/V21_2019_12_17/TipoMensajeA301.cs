using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "MensajeAlta", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    public class TipoMensajeA301
    {
        public Cabecera Cabecera { get; set; }
        public Alta Alta { get; set; }

        public TipoMensajeA301()
        {
            Cabecera = new Cabecera();
            Alta = new Alta();

        }
    }
}
