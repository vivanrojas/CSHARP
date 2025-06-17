using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public  class TAM_factura
    {
        public Int32 cemptitu { get; set; }
        public string ccounips { get; set; }
        public string cupree { get; set; }
        public string dapersoc { get; set; }
        public string cnifdnic { get; set; }
        public Int32 mes_hasta { get; set; }
        public double ifactura { get; set; }
        public int diaFact { get; set; }
        public DateTime ffactdes { get; set; }
        public DateTime ffacthas { get; set; }
        public DateTime ffactura { get; set; }
        public long creferen { get; set; }
        public int secfactu { get; set; }
        public string testfact { get; set; }
        public bool calculoTamEspecial { get; set; }

        public List<TAM_factura_detalle> fd { get; set; }

        public TAM_factura()
        {
            calculoTamEspecial = false;
            fd = new List<TAM_factura_detalle>();
        }
    }
}
