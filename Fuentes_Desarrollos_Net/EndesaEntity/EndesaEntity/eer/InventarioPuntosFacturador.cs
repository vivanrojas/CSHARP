using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.eer
{
    public class InventarioPuntosFacturador
    {
        public string nif { get; set; }
        public string cliente { get; set; }
        public string cups20 { get; set; }

        public DateTime fecha_consumo_desde { get; set; }
        public DateTime fecha_consumo_hasta { get; set; }
        public string tarifa { get; set; }
        public string num_factura { get; set; }

        public DateTime fecha_factura { get; set; }
        public string aviso { get; set; }
        public string estado { get; set; }

        public int tipo_punto_medida { get; set; }


    }
}
