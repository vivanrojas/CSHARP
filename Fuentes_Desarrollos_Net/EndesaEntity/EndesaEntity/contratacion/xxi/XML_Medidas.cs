using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.xxi
{
    public class XML_Medidas
    {
        public string tipoDHEdM { get; set; }
        public string periodo { get; set; }
        public string magnitudMedida { get; set; }
        public string procedencia { get; set; }
        public string ultimaLecturaFirme { get; set; }
        public DateTime fechaLecturaFirme { get; set; }
    }
}
