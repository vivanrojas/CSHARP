using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Autoconsumo")]
    public class AutoconsumoSolicitudAlta
    {
       
        public DatosSuministroSolicitud DatosSuministro { get; set; }
        public DatosCAUAlta DatosCAU { get; set; }

        public AutoconsumoSolicitudAlta()
        {
            DatosSuministro = new DatosSuministroSolicitud();
            DatosCAU = new DatosCAUAlta();
        }
    }

}
