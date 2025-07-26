using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "CondicionesContractuales")]
    public class CondicionesContractuales_C105
    {
        [XmlElement(ElementName = "TarifaATR")] public string TarifaATR { get; set; }
        [XmlElement(ElementName = "PeriodicidadFacturacion")] public string PeriodicidadFacturacion { get; set; }
        [XmlElement(ElementName = "TipodeTelegestion")] public string TipoTelegestion { get; set; }        
        public PotenciasContratadas PotenciasContratadas { get; set; }
        //[XmlElement(ElementName = "ModoControlPotencia")] public string ModoControlPotencia { get; set; }
       // [XmlElement(ElementName = "MarcaMedidaConPerdidas")] public string MarcaMedidaConPerdidas { get; set; }
        [XmlElement(ElementName = "TensionDelSuministro")] public string TensionDelSuministro { get; set; }
        //[XmlElement(ElementName = "VAsTrafo")] public string VAsTrafo { get; set; }

        

        public CondicionesContractuales_C105()
        {
            PotenciasContratadas = new PotenciasContratadas();
        }
    }
}
