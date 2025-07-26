using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class Condiciones_pago_tabla
    {
        public int cemptitu { get; set; }
        public int clinneg { get; set; }
        public long contrext { get; set; }
        public long ccontrps { get; set; }
        public int versi { get; set; }
        public int testcont { get; set; }
        public int ccuenta { get; set; }
        public DateTime fprealta { get; set; }
        public DateTime faltacon { get; set; }
        public DateTime fpsercon { get; set; }
        public DateTime fprevbaj { get; set; }
        public DateTime fbaja { get; set; }
        public int ccliente { get; set; }
        public DateTime ffinvigr { get; set; }
        public DateTime faltacta { get; set; }
        public DateTime fbjacta { get; set; }
        public string tmodopta { get; set; }
        public int cmedpago { get; set; }
        public string csucursa { get; set; }
        public string cctacorr { get; set; }
        public string cdigcont { get; set; }
        public int tforacor { get; set; }
        public int vfrpdcob { get; set; }
        public string tfacscon { get; set; }
        public string tultdpag { get; set; }
        public int vnumplaz { get; set; }
        public int vpridpag { get; set; }
        public int vsegdpag { get; set; }
        public int ddevengo { get; set; }
        public string tgespers { get; set; }
        public string tentrcl1 { get; set; }
        public string tindemit { get; set; }
        public bool tagrupac { get; set; }
        public string taltapre { get; set; }
        public string taplptac { get; set; }
        public string tapcliesp { get; set; }
        public string tcontsva { get; set; }
        public string csegmerc { get; set; }
        public List<Condiciones_pago_detalle_tabla> ld { get; set; }

        public Condiciones_pago_tabla()
        {
            ld = new List<Condiciones_pago_detalle_tabla>();
        }
    }
}
