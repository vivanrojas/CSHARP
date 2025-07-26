using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.gas
{
    public class SolicitudDetalle
    {
        public int id { get; set; }
        public int linea { get; set; }
        public int tipo_producto { get; set; }
        public string codigo_producto { get; set; }
        public string producto { get; set; }
        public double qd { get; set; }
        public double qa { get; set; }
        public double qi { get; set; }
        public DateTime fecha_inicio { get; set; }
        public DateTime fecha_fin { get; set; }
        public DateTime hora_inicio { get; set; }
        public string tarifa { get; set; }
    }
}
