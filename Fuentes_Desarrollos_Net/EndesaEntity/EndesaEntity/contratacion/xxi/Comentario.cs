using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.xxi
{
    public class Comentario: Audit
    {
        
        public string id_comentario { get; set; }
        public int linea { get; set; }
        public string comentario { get; set; }           
    }
}
