using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "AceptacionCambiodeComercializadorSinCambios")]
    public class AceptacionCambiodeComercializadorSinCambios
    {
        public DatosAceptacion DatosAceptacion { get; set; }
       
        public Contrato_C102_A Contrato { get; set; }
        //public ContratoAlta Contrato { get; set; }

        public AceptacionCambiodeComercializadorSinCambios()
        {
            DatosAceptacion = new DatosAceptacion();
            Contrato = new Contrato_C102_A();
        }
    }
}
