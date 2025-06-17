using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.cola
{
    public class ProcesoCola
    {
        public string cola { get; set; }
        public int grupo { get; set; }
        public int proceso { get; set; }
        public int id_p_proceso { get; set; }
        public  string ruta { get; set; }
        public string bbdd { get; set; }
        public string nombre_proceso {get;set;}
        public bool activo { get; set; }
        public bool obligatorio { get; set; }
        public  bool una_vez { get; set; }
        public string descripcion { get; set; }
        public DateTime fecha_ultima_ejec_ok { get; set; }
        public DateTime fecha_ultimo_lanzamiento { get; set; }
        public int id_p_periodicidad { get; set; }
        public int id_p_parametro { get; set; }
        public string parametro { get; set; }
        public string mensaje_error { get; set; }

    }
}
