using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.herramientas
{
    public class Seguimiento_Procesos
    {
        public string prefijo_archivo { get; set; }
        public string extractor { get; set; }
        public string area { get; set; }
        public string proceso { get; set; }        
        public string descripcion { get; set; }
        public DateTime fecha_inicio { get; set; }
        public DateTime fecha_fin { get; set; }
        public DateTime hora_inicio { get; set; }
        public string ejecucion { get; set; }
        public string tarea { get; set; }
        public string contacto { get; set; }
        public string comentario { get; set; }
        public List< Seguimiento_ProcesosDetalle> detalle { get; set; }
        public Seguimiento_Procesos()
        {
            detalle = new List<Seguimiento_ProcesosDetalle>();
        }
    }
}
