using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion
{
    public class ContratosPS_Tabla
    {
        public string cups13 { get; set; }
        public string cups20 { get; set; }
        public DateTime fecha_alta_contrato { get; set; }
        public DateTime fecha_puesta_en_servicio { get; set; }
        public  int version_contrato { get; set; }
        public int tipo_estado_contrato { get; set; }
        public DateTime fecha_inicio_version { get; set; }
        
    }
}
