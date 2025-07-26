using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V30_2022_21_01
{
    [XmlRoot(ElementName = "Contrato")]
    public class Contrato_M1
    {
        [XmlElement(ElementName = "FechaFinalizacion")] public string FechaFinalizacion { get; set; }
        public EndesaEntity.cnmc.V21_2019_12_17.AutoconsumoSolicitudAlta Autoconsumo {  get; set; }
        //[XmlElement(ElementName = "TipoAutoconsumo")] public string TipoAutoconsumo { get; set; }
        [XmlElement(ElementName = "TipoContratoATR")] public string TipoContratoATR { get; set; }
        [XmlElement(ElementName = "CUPSPrincipal")] public string CUPSPrincipal { get; set; }

        public EndesaEntity.cnmc.V21_2019_12_17.CondicionesContractuales CondicionesContractuales { get; set; }

        public EndesaEntity.cnmc.V21_2019_12_17.Contacto Contacto { get; set; }

        

        public Contrato_M1()  // irh _ constructor -  inicializa los objetos anidados de esa forma no estan en null ...
        {
            Autoconsumo = new V21_2019_12_17.AutoconsumoSolicitudAlta();
            CondicionesContractuales = new V21_2019_12_17.CondicionesContractuales();
        
        }

    }
}
