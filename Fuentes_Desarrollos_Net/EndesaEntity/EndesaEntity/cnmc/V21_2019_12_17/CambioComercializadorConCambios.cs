using EndesaEntity.cnmc.V30_2022_21_01;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "CambiodeComercializadorConCambios")]
    public class CambioComercializadorConCambios
    {
        public DatosSolicitud_C1 DatosSolicitud { get; set; }
        public V30_2022_21_01.Contrato_C201 Contrato { get; set; }
        public Cliente Cliente { get; set; }
        public RegistrosDocumento RegistrosDocumento { get; set; }

        public CambioComercializadorConCambios()
        {
            DatosSolicitud = new DatosSolicitud_C1();
         //   Contrato = new Contrato_C201();  //  inicializado
        }

    }
}
