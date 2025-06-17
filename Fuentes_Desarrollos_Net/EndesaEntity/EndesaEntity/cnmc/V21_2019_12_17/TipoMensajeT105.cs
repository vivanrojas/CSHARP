using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "MensajeActivacionTraspasoCOR")]
    public class TipoMensajeT105
    {
                
        public Cabecera Cabecera { get; set; }
        public ActivacionTraspasoCOR ActivacionTraspasoCOR { get; set; }
       
        public TipoMensajeT105()
        {
            Cabecera = new Cabecera();
            ActivacionTraspasoCOR = new ActivacionTraspasoCOR();
        }

    }
}
