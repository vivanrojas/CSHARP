using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Alta")]
    public class Alta
    {
        public DatosSolicitud_A3 DatosSolicitud { get; set; }
        public Contrato Contrato { get; set; }
        public ClienteConDireccion Cliente { get; set; }    
        public RegistrosDocumento RegistrosDocumento { get; set; }
        public Alta()
        {
            DatosSolicitud = new DatosSolicitud_A3();
            Contrato = new Contrato();
            Cliente = new ClienteConDireccion();
            RegistrosDocumento = new RegistrosDocumento();

        }
    }
}
