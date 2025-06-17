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
    //[XmlRoot(ElementName = "ModificacionDeATR", Namespace = "http://localhost/elegibilidad:ModificacionDeATR", IsNullable = true)]
    [XmlRoot(ElementName = "MensajeModificacionDeATR", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    
    public class TipoMensajeM101
    {
        public Cabecera Cabecera { get; set; }
        public ModificacionDeATR ModificacionDeATR { get; set; }

        public TipoMensajeM101()
        {
            Cabecera = new Cabecera();
            ModificacionDeATR = new ModificacionDeATR();

        }



    }
}

