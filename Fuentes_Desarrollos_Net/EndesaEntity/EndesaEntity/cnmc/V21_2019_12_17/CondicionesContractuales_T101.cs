using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "CondicionesContractuales")]
    public class CondicionesContractuales_T101
    {
        [XmlElement(ElementName = "TarifaATR")] public string TarifaATR { get; set; }

        public PotenciasContratadas PotenciasContratadas { get; set; }

        [XmlElement(ElementName = "PeriodicidadFacturacion")] public string PeriodicidadFacturacion { get; set; }
        [XmlElement(ElementName = "ConsumoAnualEstimado")] public string ConsumoAnualEstimado { get; set; }

        public CondicionesContractuales_T101()
        {
            PotenciasContratadas = new PotenciasContratadas();
        }
    }
}
