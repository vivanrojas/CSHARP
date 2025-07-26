using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class CalendarioTarifario
    {
        public string tarifa { get; set; }
        public string territorio { get; set; }
        public List<FechaTarifa> ft { get; set; }

        public CalendarioTarifario()
        {
            ft = new List<FechaTarifa>();

        }
    }
}
