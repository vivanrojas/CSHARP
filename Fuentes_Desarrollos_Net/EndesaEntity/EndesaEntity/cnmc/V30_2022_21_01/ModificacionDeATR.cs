using EndesaEntity.cnmc.V21_2019_12_17;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V30_2022_21_01
{
    [XmlRoot(ElementName = "ModificacionDeATR")]
    public class ModificacionDeATR
    {
        public DatosSolicitud_M1 DatosSolicitud { get; set; }
        public Contrato_M1 Contrato { get; set; }
        public EndesaEntity.cnmc.V21_2019_12_17.Cliente Cliente { get; set; }
        public RegistrosDocumento RegistrosDocumento { get; set; }

        public ModificacionDeATR()
        {
            DatosSolicitud = new DatosSolicitud_M1();
            Contrato = new Contrato_M1();
            Cliente = new V21_2019_12_17.Cliente();
            RegistrosDocumento = new RegistrosDocumento();
        }

    }
}
