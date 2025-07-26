using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion
{
    public class Documento_Tabla
    {
        public int id_inventario { get; set; }
        public string codigosolicitud { get; set; }
        public int id_estado_documento {get;set;}
        public DateTime fecha { get; set; }
        public string observaciones { get; set; }

    }
}
