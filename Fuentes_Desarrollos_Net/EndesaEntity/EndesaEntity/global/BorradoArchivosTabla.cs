using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.global
{
    public class BorradoArchivosTabla: Audit
    {
        public string prefijo_archivo { get; set; }
        public DateTime from_date { get; set; }
        public DateTime to_date { get; set; }
        public string formato_fecha { get; set; }
        public string ubicacion { get; set; }
        public string accion { get; set; }
        public string destino { get; set; }
        public bool conservar_fin_mes { get; set; }
        public int numero_ultimos_archivos_a_conservar { get; set; }
        public string descripcion { get; set; }


    }
}
