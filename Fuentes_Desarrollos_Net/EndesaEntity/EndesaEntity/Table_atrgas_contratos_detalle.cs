using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Table_atrgas_contratos_detalle : FieldsAudit
    {
        public string nif { get; set; }
        public string cups20 { get; set; }
        public DateTime fecha_inicio { get; set; }
        public DateTime fecha_fin { get; set; }
        public double qd { get; set; }
        public double qi { get; set; }
        public DateTime hora_inicio { get; set; }
        public string tarifa { get; set; }
        public string tipo { get; set; }
        public bool nuevo_registro { get; set; }
        public string comentario { get; set; }
        public long id_solicitud { get; set; }
        public int linea_solicitud { get; set; }
    }
}
