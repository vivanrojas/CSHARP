using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "BajaSuspension")]
    public class BajaSuspension
    {
        public DatosSolicitud DatosSolicitud { get; set; }
        public Cliente Cliente { get; set; }
        public Contacto Contacto { get; set; }
        public RegistrosDocumento RegistrosDocumento { get; set; }

        public BajaSuspension()
        {
            DatosSolicitud = new DatosSolicitud();
            Cliente = new Cliente();
            Contacto = new Contacto();
            RegistrosDocumento = new RegistrosDocumento();
        }


    }
}
