using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.xxi
{
    public class Casos_Tabla
    {
        public int caso { get; set; }
        public string tipo { get; set; }
        public bool existe_baja { get; set; }
        public bool existe_ps { get; set; }
        public string empresa_ps { get; set; }
        public int estado_contrato_ps_id { get; set; }
        public string estado_contrato_ps_descripcion {get;set;}
        public bool crear_contrato { get; set; }
        public bool crear_incidencia { get; set; }
        public bool realizar_seguimiento { get; set; }
        public string acciones { get; set; }
        public string observaciones { get; set; }
    }
}
