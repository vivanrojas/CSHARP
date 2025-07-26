using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;


namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "AceptacionBajaSuspension", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    public class AceptacionBajaSuspension
    {
        public DatosAceptacionB102  DatosAceptacion { get; set; }

        // public ContratoAlta Contrato { get; set; }

        public AceptacionBajaSuspension()
        {
           DatosAceptacion = new DatosAceptacionB102();

            //   Contrato = new ContratoAlta();
        }
    }
}
