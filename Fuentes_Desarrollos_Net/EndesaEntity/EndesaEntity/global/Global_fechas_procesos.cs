using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.global
{
    public class Global_fechas_procesos
    {
        public string area { get; set; }
        public string proceso { get; set; }
        public string paso { get; set; }
        public string descripcion { get; set; }
        public DateTime fecha { get; set; }
        public DateTime fecha_ultima_ejecucion { get; set; }
        public DateTime fecha_inicio { get; set; }
        public DateTime fecha_fin { get; set; }
        public DateTime f_ult_mod { get; set; }
    }
}
