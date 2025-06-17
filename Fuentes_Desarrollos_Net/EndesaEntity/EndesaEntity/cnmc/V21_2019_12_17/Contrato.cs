//using EndesaEntity.facturacion.cuadroDeMando;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Contrato")]
    public class Contrato
    {
        public IdContrato IdContrato { get; set; }

        [XmlElement(ElementName = "FechaFinalizacion")] public string FechaFinalizacion { get; set; }    
        [XmlElement(ElementName = "TipoAutoconsumo")] public string TipoAutoconsumo { get; set; }
       
        [XmlElement(ElementName = "TipoContratoATR")] public string TipoContratoATR { get; set; }
        public CondicionesContractuales CondicionesContractuales { get; set; }
        
        //Datos usados en caso de aceptación (A302-A)
        [XmlElement(ElementName = "TipoActivacionPrevista")] public string TipoActivacionPrevista { get; set; }
        [XmlElement(ElementName = "FechaActivacionPrevista")] public string FechaActivacionPrevista { get; set; }
        
        public Contacto Contacto { get; set; }
        
        public Contrato()
        {
            
            CondicionesContractuales = new CondicionesContractuales();

        }
    }
}
