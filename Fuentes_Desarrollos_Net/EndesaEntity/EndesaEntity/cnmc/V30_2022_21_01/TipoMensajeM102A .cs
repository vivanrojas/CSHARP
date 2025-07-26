using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using EndesaEntity.cnmc.V21_2019_12_17;

namespace EndesaEntity.cnmc.V30_2022_21_01
{
    [XmlRoot(ElementName = "MensajeAceptacionModificacionDeATR", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    public class TipoMensajeM102A
    {
        public Cabecera Cabecera { get; set; }
        public AceptacionModificionDeATR AceptacionModificionDeATR { get; set; }

        public TipoMensajeM102A()
        {
            Cabecera = new Cabecera();
            AceptacionModificionDeATR = new AceptacionModificionDeATR();
        }
    }
}
