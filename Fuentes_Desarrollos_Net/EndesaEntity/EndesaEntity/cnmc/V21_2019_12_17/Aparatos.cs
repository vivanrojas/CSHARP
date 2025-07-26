using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Aparatos")]
    public class Aparatos
    {
        List<Aparato> ListaAparato { get; set; }
        [XmlElement(ElementName = "TipoMovimiento")] public string TipoMovimiento { get; set; }
        [XmlElement(ElementName = "TipoEquipoMedida")] public string TipoEquipoMedida { get; set; }
        [XmlElement(ElementName = "TipoPropiedadAparato")] public string TipoPropiedadAparato { get; set; }
        [XmlElement(ElementName = "TipoDHEdM")] public string TipoDHEdM { get; set; }
        [XmlElement(ElementName = "ModoMedidaPotencia")] public string ModoMedidaPotencia { get; set; }
        [XmlElement(ElementName = "CodPrecinto")] public string CodPrecinto { get; set; }

    }
}
