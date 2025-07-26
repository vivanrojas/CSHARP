using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "MensajeAceptacionAlta", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    public class TipoMensajeC202_A
    {
        public Cabecera Cabecera { get; set; }
        public AceptacionAlta AceptacionAlta { get; set; }

        public TipoMensajeC202_A()
        {
            Cabecera = new Cabecera();
            AceptacionAlta = new AceptacionAlta();
        }
    }
}
