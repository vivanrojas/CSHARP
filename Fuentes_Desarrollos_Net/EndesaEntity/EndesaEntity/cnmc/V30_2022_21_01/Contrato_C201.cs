using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using EndesaEntity.cnmc.V21_2019_12_17;
using EndesaEntity.cnmc.V30_2022_21_01;

namespace EndesaEntity.cnmc.V30_2022_21_01  // cambio de version V21_2019_12_17
{
    [XmlRoot(ElementName = "Contrato")]
    public class Contrato_C201
   

    {
        [XmlElement(ElementName = "FechaFinalizacion")] public string FechaFinalizacion { get; set; }



         [XmlElement(ElementName = "Autoconsumo", Namespace = "http://localhost/elegibilidad")]

         public V21_2019_12_17.AutoconsumoSolicitudAlta Autoconsumo { get; set; }//

        [XmlElement(ElementName = "TipoContratoATR")] public string TipoContratoATR { get; set; }
        [XmlElement(ElementName = "CUPSPrincipal")] public string CUPSPrincipal { get; set; }

        
        public V30_2022_21_01.CondicionesContractuales_C201 CondicionesContractuales { get; set; }
           
             
        public V21_2019_12_17.Contacto Contacto { get; set; }
        public Contrato_C201()
        {
            //Autoconsumo = new AutoconsumoSolicitudAlta();//
            //CondicionesContractuales = new CondicionesContractuales_C201();
           // Contacto = new Contacto();//

        }
    }
}
