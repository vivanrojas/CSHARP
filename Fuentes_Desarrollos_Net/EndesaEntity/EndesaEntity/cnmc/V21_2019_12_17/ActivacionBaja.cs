using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
 
    [XmlRoot(ElementName = "ActivacionBaja", Namespace = "http://localhost/elegibilidad")]
    [XmlType(Namespace = "http://localhost/elegibilidad")]
    public class ActivacionBaja
    { 
        public DatosActivacionBaja DatosActivacionBaja { get; set; }

        public Contrato_B105 Contrato { get; set; }
        
        public PuntosDeMedida PuntosDeMedida { get; set; }
        public ActivacionBaja()
        {
            DatosActivacionBaja = new DatosActivacionBaja();

            Contrato = new Contrato_B105();

            PuntosDeMedida = new PuntosDeMedida();
        }

    }
}
