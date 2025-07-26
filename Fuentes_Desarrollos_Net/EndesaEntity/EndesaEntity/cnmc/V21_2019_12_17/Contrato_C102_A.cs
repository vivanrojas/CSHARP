using EndesaEntity.facturacion.cuadroDeMando;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Contrato")]
    public class Contrato_C102_A
    {
       // public IdContrato IdContrato { get; set; }
      //  [XmlElement(ElementName = "FechaFinalizacion")] public string FechaFinalizacion { get; set; }
        
      //  public AutoconsumoSolicitudAlta Autoconsumo { get; set; }

        //Debe incluirse dentro del objeto Autoconsumo
        //[XmlElement(ElementName = "TipoAutoconsumo")] public string TipoAutoconsumo { get; set; }

        [XmlElement(ElementName = "TipoContratoATR")] public string TipoContratoATR { get; set; }
     //   [XmlElement(ElementName = "CUPSPrincipal")] public string CUPSPrincipal { get; set; }
        public CondicionesContractuales CondicionesContractuales { get; set; }
     //   [XmlElement(ElementName = "ConsumoAnualEstimado")] public string ConsumoAnualEstimado { get; set; }
    //    public Contacto Contacto { get; set; }

        //Datos usados en caso de aceptación (A302-A)
        [XmlElement(ElementName = "TipoActivacionPrevista")] public string TipoActivacionPrevista { get; set; }
        [XmlElement(ElementName = "FechaActivacionPrevista")] public string FechaActivacionPrevista { get; set; }
        
        public Contrato_C102_A()
        {

            CondicionesContractuales = new CondicionesContractuales();
            //NO ES OBLIGATORIO
            //Autoconsumo = new AutoconsumoSolicitudAlta();
           // Contacto = new Contacto();

        }
    }
}
