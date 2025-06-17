using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.xxi
{
    public class Casos_Tabla_Historico : Casos_Tabla
    {
        public string nif { get; set; }
        public string razon_social { get; set; }
        public string tarifa { get; set; }
        public string cups { get; set; }
        public string codigo_solicitud { get; set; }
        public DateTime fecha_alta { get; set; }
        public DateTime fecha_baja { get; set; }
        public bool existe_alta { get; set; }
        public string estado_contrato_ps { get; set; }
        public bool documentacion_impresa { get; set; }
        public int estado_id { get; set; }
        public string descripcion_estado_caso { get; set; }
        public string comentario { get; set; }

    }
}
