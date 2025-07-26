using EndesaEntity.cnmc.V21_2019_12_17;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V30_2022_21_01
{
    [XmlRoot(ElementName = "ActivacionAlta")]
    class ActivacionAltaA3
    {

        public DatosActivacion DatosActivacion { get; set; }
        public ContratoAlta Contrato { get; set; }
        public PuntosDeMedida PuntosDeMedida { get; set; }
        public ActivacionAltaA3()
        {
            DatosActivacion = new DatosActivacion();
            Contrato = new ContratoAlta();
            PuntosDeMedida = new PuntosDeMedida();
        }

    }
}
