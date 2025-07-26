using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{

    [XmlRoot(ElementName = "MensajeActivacionAlta", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    //  [XmlRoot(ElementName = "MensajeAceptacionAlta", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    public class TipoMensajeA305
    {
        public Cabecera Cabecera { get; set; }

        public ActivacionAlta ActivacionAlta { get; set; }

        public TipoMensajeA305()
        {
            Cabecera = new Cabecera();
            // irh
           // ActivacionAlta = new ActivacionAlta();
        }
    }
}
