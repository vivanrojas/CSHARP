using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "CondicionesContractuales")]
    public class CondicionesContractuales_M102
    {
        [XmlElement(ElementName = "TarifaATR")] public string TarifaATR { get; set; }

        public PotenciasContratadas PotenciasContratadas { get; set; }

        [XmlElement(ElementName = "ModoControlPotencia")] public string ModoControlPotencia { get; set; }
       
        public CondicionesContractuales_M102()
        {
            PotenciasContratadas = new PotenciasContratadas();
        }
    }
}
