using EndesaEntity.cnmc.V30_2022_21_01;
using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "MensajeAceptacionBajaSuspension", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    public class TipoMensajeB102_A
    {
        public Cabecera Cabecera { get; set; }
       // public AceptacionAlta AceptacionAlta { get; set; }
        public AceptacionBajaSuspension AceptacionBajaSuspension { get; set; }  
        public TipoMensajeB102_A()
        {
            Cabecera = new Cabecera();
            AceptacionBajaSuspension = new AceptacionBajaSuspension();
        }
    }
}
