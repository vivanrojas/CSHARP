using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.xxi
{
    public class Eventuales : XML_Datos
    {
        public string distribuidora { get; set; }
        public string indActivacion { get; set; }

        public string contacto { get; set; }
        public DateTime fecha_de_baja { get; set; }
        public string tipoDocAportado { get; set; }
        public string direccionURL { get; set; }
        public string observaciones { get; set; }
    }
    
}
