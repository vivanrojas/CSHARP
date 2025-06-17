using EndesaEntity.cnmc.V21_2019_12_17;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V30_2022_21_01
{
    [XmlRoot(ElementName = "Cliente")]
    public class Cliente
    {
        public IdCliente IdCliente { get; set; }
        public Nombre Nombre { get; set; }
        public Telefono Telefono { get; set; }
        public Direccion Direccion { get; set; }
        [XmlElement(ElementName = "IndicadorTipoDireccion")] public string IndicadorTipoDireccion { get; set; }
        public Cliente()
        {
            IdCliente = new IdCliente();
            Nombre = new Nombre();
            Telefono = new Telefono();
            

        }
    }
}
