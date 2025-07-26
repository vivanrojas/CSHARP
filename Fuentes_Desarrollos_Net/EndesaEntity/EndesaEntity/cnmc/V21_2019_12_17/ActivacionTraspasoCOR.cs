using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "ActivacionTraspasoCOR")]
    public class ActivacionTraspasoCOR
    {
        public DatosActivacion DatosActivacion { get; set; }
        public Contrato Contrato { get; set; }
        public PuntosDeMedida PuntosDeMedida { get; set; }



        public ActivacionTraspasoCOR()
        {
            DatosActivacion = new DatosActivacion();
            Contrato = new Contrato(); 
            PuntosDeMedida = new PuntosDeMedida();
        }
    }
}
