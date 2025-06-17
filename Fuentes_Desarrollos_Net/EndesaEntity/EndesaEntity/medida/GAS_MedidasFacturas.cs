using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class GAS_MedidasFacturas
    {
        public int id_PS { get; set; }
        public int mes { get; set; }
        public string comentario { get; set; }
        public bool facturado { get; set; }
        public bool existe { get; set; }
        public DateTime medida { get; set; }

    }
}
