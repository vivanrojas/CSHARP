using EndesaEntity.facturacion.cuadroDeMando;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
  // [XmlRoot(ElementName = "Contrato")]
    [XmlRoot(ElementName = "Contrato", Namespace = "http://localhost/elegibilidad")]
    [XmlType(Namespace = "http://localhost/elegibilidad")]
    public class Contrato_C105
    {
        [XmlElement(ElementName = "IdContrato", Namespace = "http://localhost/elegibilidad")]
        public IdContrato IdContrato { get; set; }

        [XmlElement(ElementName = "FechaFinalizacion", Namespace = "http://localhost/elegibilidad")]
        public string FechaFinalizacion { get; set; }
        

        public AutoconsumoSolicitudAlta Autoconsumo { get; set; }

        [XmlElement(ElementName = "TipoContratoATR", Namespace = "http://localhost/elegibilidad")]
        public string TipoContratoATR { get; set; }

        //CUPSPrincipal

        [XmlElement(ElementName = "CondicionesContractuales", Namespace = "http://localhost/elegibilidad")]
        public CondicionesContractuales CondicionesContractuales { get; set; }
            
                      
        public Contrato_C105()
        {

            CondicionesContractuales = new CondicionesContractuales();
           
        }
    }
}
