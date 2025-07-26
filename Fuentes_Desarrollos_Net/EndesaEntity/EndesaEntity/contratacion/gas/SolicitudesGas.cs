using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.gas
{
    public class SolicitudesGas
    {


        public long id { get; set; }
        public int linea { get; set; }
        public string cnifdnic { get; set; }
        public string dapersoc { get; set; }
        public string distribuidora { get; set; }
        public string remitente { get; set; }
        public DateTime fecha_mail { get; set; }
        public string cups { get; set; }
        public string producto { get; set; }
        public double qd { get; set; }
        public DateTime fecha_inicio { get; set; }
        public DateTime fecha_fin { get; set; }
        public double qi { get; set; }
        public DateTime hora_inicio { get; set; }
        public string tarifa { get; set; }
        
    }
}
