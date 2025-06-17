using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class Fact_pt_calendarios_Tabla
    {
        public int dia { get; set; }
        public string estacion { get; set; }
        public int periodo { get; set; }
        public int hora_desde { get; set; }
        public int hora_hasta { get; set; }
    }
}
