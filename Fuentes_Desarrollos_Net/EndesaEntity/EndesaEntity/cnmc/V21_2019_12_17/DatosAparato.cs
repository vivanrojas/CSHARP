using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{

    [XmlRoot(ElementName = "DatosAparato")]
    public class DatosAparato
    {
        [XmlElement(ElementName = "PeriodoFabricacion")] public string PeriodoFabricacion { get; set; }
        [XmlElement(ElementName = "NumeroSerie")] public string NumeroSerie { get; set; }
        [XmlElement(ElementName = "FuncionAparato")] public string FuncionAparato { get; set; }
        [XmlElement(ElementName = "NumIntegradores")] public string NumIntegradores { get; set; }
        [XmlElement(ElementName = "ConstanteEnergia")] public string ConstanteEnergia { get; set; }
        [XmlElement(ElementName = "ConstanteMaximetro")] public string ConstanteMaximetro { get; set; }
        [XmlElement(ElementName = "RuedasEnteras")] public string RuedasEnteras { get; set; }
        [XmlElement(ElementName = "RuedasDecimales")] public string RuedasDecimales { get; set; }
    }
}
