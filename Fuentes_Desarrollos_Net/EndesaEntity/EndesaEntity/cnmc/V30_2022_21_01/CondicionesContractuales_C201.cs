using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using EndesaEntity.cnmc.V21_2019_12_17;
using EndesaEntity.cnmc.V30_2022_21_01;

namespace EndesaEntity.cnmc.V30_2022_21_01
{
    [XmlRoot(ElementName = "CondicionesContractuales")]
    public class CondicionesContractuales_C201
    {
        [XmlElement(ElementName = "TarifaATR")] public string TarifaATR { get; set; }

        public PotenciasContratadas PotenciasContratadas { get; set; }

        [XmlElement(ElementName = "PeriodicidadFacturacion")] public string PeriodicidadFacturacion { get; set; }
       // [XmlElement(ElementName = "ConsumoAnualEstimado")] public string ConsumoAnualEstimado { get; set; }


        public CondicionesContractuales_C201()
        {
            PotenciasContratadas = new PotenciasContratadas();
        }

    }
}
