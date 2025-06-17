using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class Anomalia
    {
        public string cups20 { get; set; }
        public string[] anomalia { get; set; }

        public bool existe { get; set; }

        public Anomalia()
        {
            anomalia = new string[5];
        }
    }
}
