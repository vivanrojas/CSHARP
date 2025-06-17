using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class PO1011_Informe
    {
        public int cups_id { get; set; }
        public DateTime fecha_hora { get; set; }
        public double activa { get; set; }
        public double reactiva { get; set; }
        public int total_periodos { get; set; }
        public bool completo { get; set; }

    }
}
