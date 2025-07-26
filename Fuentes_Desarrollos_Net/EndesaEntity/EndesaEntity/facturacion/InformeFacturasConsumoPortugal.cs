using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class InformeFacturasConsumoPortugal
    {
        public string empresa { get; set; }
        public string nif { get; set; }
        public string razon_social { get; set; }
        public string cupsree { get; set; }
        public long creferen { get; set; }
        public int secfactu { get; set; }
        public string testfact { get; set; }

        public string cfactura { get; set; }
        public DateTime ffactura { get; set; }
        public DateTime ffactdes { get; set; }
        public DateTime ffacthas { get; set; }

        public int vcuovafa { get; set; }
        public int p1 { get; set; }
        public int p2 { get; set; }
        public int p3 { get; set; }
        public int p4 { get; set; }
        public int p5 { get; set; }
        public int p6 { get; set; }
        public DateTime min_ffactdes { get; set; }
        public DateTime max_ffacthas { get; set; }        

        public InformeFacturasConsumoPortugal()
        {
            vcuovafa = 0;
            p1 = 0;
            p2 = 0;
            p3 = 0;
            p4 = 0;
               
        }

    }
}
