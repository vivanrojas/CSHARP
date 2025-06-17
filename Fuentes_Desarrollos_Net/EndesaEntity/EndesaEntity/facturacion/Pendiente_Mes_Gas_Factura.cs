using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class Pendiente_Mes_Gas_Factura
    {
        public string cups { get; set; }
        public int id_ps { get; set; }
        public string cif { get; set; }
        public string cliente { get; set; }
        public string atr { get; set; }
        public bool top { get; set; }
        public  int mes { get; set; }
        public DateTime medida { get; set; }
        public string comentario { get; set; }
        public string facturacion { get; set; }
        public string distribuidora { get; set; }
        public DateTime carga { get; set; }

    }
}
