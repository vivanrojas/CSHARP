using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{

    [XmlRoot(ElementName = "MensajeActivacionBajaSuspension", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
       public class TipoMensajeB105
    {
        public Cabecera Cabecera { get; set; }

        public ActivacionBaja ActivacionBaja { get; set; }

        public TipoMensajeB105()
        {
            Cabecera = new Cabecera();
           
            ActivacionBaja = new ActivacionBaja();
        }
    }
}
