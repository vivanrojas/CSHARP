using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Contrato")]
    public class Contrato_T101
    {
        
        [XmlElement(ElementName = "FechaFinalizacion")] public string FechaFinalizacion { get; set; }
        [XmlElement(ElementName = "TipoAutoconsumo")] public string TipoAutoconsumo { get; set; }
        [XmlElement(ElementName = "TipoContratoATR")] public string TipoContratoATR { get; set; }

        public CondicionesContractuales_T101 CondicionesContractuales { get; set; }

        public Contacto Contacto { get; set; }
        public Contrato_T101()
        {
            
            CondicionesContractuales = new CondicionesContractuales_T101();
            Contacto = new Contacto();

        }
    }
}
