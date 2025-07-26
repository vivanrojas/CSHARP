using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    //[XmlRoot(ElementName = "ActivacionAlta")]
    [XmlRoot(ElementName = "ActivacionAlta", Namespace = "http://localhost/elegibilidad")]
    [XmlType(Namespace = "http://localhost/elegibilidad")]
    public class ActivacionAlta
    {

        public DatosActivacion DatosActivacion { get; set; }
        public ContratoAlta Contrato { get; set; }
        public PuntosDeMedida PuntosDeMedida { get; set; }
        public ActivacionAlta()
        {
            DatosActivacion = new DatosActivacion();
            Contrato = new ContratoAlta();
            PuntosDeMedida = new PuntosDeMedida();
        }

    }
}
