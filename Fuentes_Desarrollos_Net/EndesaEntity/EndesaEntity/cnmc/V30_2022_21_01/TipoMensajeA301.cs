using EndesaEntity.cnmc.V21_2019_12_17;
using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V30_2022_21_01
{
    [XmlRoot(ElementName = "MensajeAlta", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    public class TipoMensajeA301
    {
        public EndesaEntity.cnmc.V21_2019_12_17.Cabecera Cabecera { get; set; }
        public AltaA301 Alta { get; set; }

        public TipoMensajeA301()
        {
            Cabecera = new EndesaEntity.cnmc.V21_2019_12_17.Cabecera();
            Alta = new AltaA301();
        }
    }
}
