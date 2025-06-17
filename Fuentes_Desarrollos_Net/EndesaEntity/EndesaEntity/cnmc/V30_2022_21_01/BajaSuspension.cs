using EndesaEntity.cnmc.V21_2019_12_17;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V30_2022_21_01
{
    [XmlRoot(ElementName = "BajaSuspension")]
    public class BajaSuspension
    {
        public DatosSolicitud_B1 DatosSolicitud { get; set; }
        public Cliente Cliente { get; set; }
        public EndesaEntity.cnmc.V21_2019_12_17.Contacto Contacto { get; set; }
        public RegistrosDocumento RegistroDocumento { get; set; }
        public BajaSuspension()
        {
            DatosSolicitud = new DatosSolicitud_B1();
            Cliente = new Cliente();
            Contacto = new EndesaEntity.cnmc.V21_2019_12_17.Contacto();            
        }
        


    }
}
