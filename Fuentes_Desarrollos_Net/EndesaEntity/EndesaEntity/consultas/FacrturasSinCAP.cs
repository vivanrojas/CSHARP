using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.consultas
{
    public class FacturasSinCAP
    {
        public string cnifdnic { get; set; }
        public string dapersoc { get; set; }
        public string ccounips { get; set; }
        public string cupsree { get; set; }
        public string ctarifa { get; set; }
        public string cfactura { get; set; }
        public DateTime ffactura { get; set; }
        public DateTime ffactdes { get; set; }
        public DateTime ffacthas { get; set; }
        public double ifactura { get; set; }        
        public string tfactura { get; set; }
        public int secfactu { get; set; }
        public long creferen { get; set; }
        public string testfact { get; set; }

        public Dictionary <int, double> dic { get; set; }

        public FacturasSinCAP()
        {
            dic = new Dictionary<int, double>();
        }
    }
}
