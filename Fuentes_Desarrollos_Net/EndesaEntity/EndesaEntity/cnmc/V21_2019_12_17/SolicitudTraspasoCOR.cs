using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "SolicitudTraspasoCOR")]
    public class SolicitudTraspasoCOR
    {
        public DatosSolicitud DatosSolicitud { get; set; }

        public Contrato_T101 Contrato { get; set; }

        public Cliente Cliente { get; set; }

        public DireccionPS DireccionPS { get; set; }
        public SolicitudTraspasoCOR()
        {
            DatosSolicitud = new DatosSolicitud();
            Contrato = new Contrato_T101();
            Cliente = new Cliente();
            DireccionPS = new DireccionPS();
        }
    }
}
