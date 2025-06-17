using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    // [XmlRoot(ElementName = "PuntoDeMedida")]
    public class PuntoDeMedida
    {
        [XmlElement(ElementName = "CodPM")] public string CodPM { get; set; }
        [XmlElement(ElementName = "TipoMovimiento")] public string TipoMovimiento { get; set; }
        [XmlElement(ElementName = "TipoPM")] public string TipoPM { get; set; }
        [XmlElement(ElementName = "ModoLectura")] public string ModoLectura { get; set; }
        [XmlElement(ElementName = "Funcion")] public string Funcion { get; set; }
        [XmlElement(ElementName = "DireccionEnlace")] public string DireccionEnlace { get; set; }
        [XmlElement(ElementName = "DireccionPuntoMedida")] public string DireccionPuntoMedida { get; set; }
        [XmlElement(ElementName = "NumLinea")] public string NumLinea { get; set; }
        [XmlElement(ElementName = "TelefonoTelemedida")] public string TelefonoTelemedida { get; set; }
        [XmlElement(ElementName = "EstadoTelefono")] public string EstadoTelefono { get; set; }
        [XmlElement(ElementName = "ClaveAcceso")] public string ClaveAcceso { get; set; }
        [XmlElement(ElementName = "TensionPM")] public string TensionPM { get; set; }
        [XmlElement(ElementName = "FechaVigor")] public string FechaVigor { get; set; }
        [XmlElement(ElementName = "FechaAlta")] public string FechaAlta { get; set; }
        public List<Aparato> Aparatos { get; set; }
        public List<Medida> Medidas { get; set; }


        public PuntoDeMedida()
        {
            Aparatos = new List<Aparato>();
        }
    }
}
