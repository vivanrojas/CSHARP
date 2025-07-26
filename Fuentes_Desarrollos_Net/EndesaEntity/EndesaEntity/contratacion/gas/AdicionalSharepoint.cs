using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.gas
{
    public class AdicionalSharepoint
    {
        public int id { get; set; }
        public DateTime hora_de_inicio { get; set; }
        public DateTime hora_de_finalizacion { get; set; }
        public string correo_electronico { get; set; }
        public string nombre { get; set; }
        public  string cliente { get; set; }
        public string cups { get; set; }
        public string producto { get; set; } 
        public DateTime fecha_inicio { get; set; }
        public bool generado_XML { get; set; }
    }
}
