using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.sigame
{
    public class Factura
    {
        public string cups20 { get; set; }
        public string cfactura { get; set; }
        public string fuente { get; set; }  
        public string medida { get; set; }
        public DateTime fd { get; set; }
        public DateTime fh { get; set; }
        public DateTime ffactura { get; set; }

        public DateTime last_update_date { get; set; }

    }
}
