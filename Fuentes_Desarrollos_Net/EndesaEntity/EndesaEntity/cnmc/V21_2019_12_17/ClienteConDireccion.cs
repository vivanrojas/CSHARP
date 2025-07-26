using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "Cliente")]
    public class ClienteConDireccion
    {
        public IdCliente IdCliente { get; set; }
        public Nombre Nombre { get; set; }
        public Telefono Telefono { get; set; }
        [XmlElement(ElementName = "IndicadorTipoDireccion")] public string IndicadorTipoDireccion { get; set; }
        public Direccion Direccion { get; set; }

        public ClienteConDireccion()
        {
            IdCliente = new IdCliente();
            Nombre = new Nombre();
            Telefono = new Telefono();
            Direccion = new Direccion();

        }
       
    }
}