using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class TotalRetrasoFacturacion
    {
        public String cnifdnic { get; set; }
        public string dapersoc { get; set; }
        public DateTime fd { get; set; }
        public DateTime fh { get; set; }
        public DateTime fa { get; set; }
        public String cups { get; set; }
        public Int32 dias { get; set; }
    }
}
