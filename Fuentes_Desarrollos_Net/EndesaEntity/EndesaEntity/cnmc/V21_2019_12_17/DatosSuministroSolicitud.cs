using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "DatosSuministro")]
    public class DatosSuministroSolicitud
    {
        [XmlElement(ElementName = "TipoCUPS")] public string TipoCUPS { get; set; }
        [XmlElement(ElementName = "RefCatastro")] public string RefCatastro { get; set; }
       

        public DatosSuministroSolicitud()
        {

        }
    }

}
