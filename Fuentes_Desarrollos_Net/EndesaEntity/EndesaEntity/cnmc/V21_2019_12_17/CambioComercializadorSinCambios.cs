using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "CambiodeComercializadorSinCambios")]
    public class CambioComercializadorSinCambios
    {
        public DatosSolicitud_C1 DatosSolicitud { get; set; }
        public Cliente Cliente { get; set; }
        public RegistrosDocumento RegistrosDocumento { get; set; }

        public CambioComercializadorSinCambios()
        {
            DatosSolicitud = new DatosSolicitud_C1();
        }

    }
}
