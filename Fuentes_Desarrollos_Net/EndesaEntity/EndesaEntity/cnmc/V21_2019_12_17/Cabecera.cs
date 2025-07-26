using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Cabecera")]
    public class Cabecera
    {

        [XmlElement(ElementName = "CodigoREEEmpresaEmisora")] public string CodigoREEEmpresaEmisora { get; set; }
        [XmlElement(ElementName = "CodigoREEEmpresaDestino")] public string CodigoREEEmpresaDestino { get; set; }
        [XmlElement(ElementName = "CodigoDelProceso")] public string CodigoDelProceso { get; set; }
        [XmlElement(ElementName = "CodigoDePaso")] public string CodigoDePaso { get; set; }
        [XmlElement(ElementName = "CodigoDeSolicitud")] public string CodigoDeSolicitud { get; set; }        
        [XmlElement(ElementName = "SecuencialDeSolicitud")] public string SecuencialDeSolicitud { get; set; }
        [XmlElement(ElementName = "FechaSolicitud")] public string FechaSolicitud { get; set; }
        [XmlElement(ElementName = "CUPS")] public string CUPS { get; set; }
    }
}
