using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class FacturaInffact
    {
        public enum linea_negocio
        {
            Luz,
            Gas
        }

        public string ccounips { get; set; }
        public string cupsree { get; set; }
        public string cemptitu { get; set; }
        public long creferen { get; set; }
        public int secfactu { get; set; }
        public string creffact { get; set; }
        public string cfactura { get; set; }
        public DateTime ffactura { get; set; }
        public DateTime ffactdes { get; set; }
        public DateTime ffacthas { get; set; }
        public double vcuovafa { get; set; }
        public string venereac { get; set; }
        public double vcuofifa { get; set; }
        public double ifactura { get; set; }
        public double iva { get; set; }
        public double iimpues2 { get; set; }
        public double iimpues3 { get; set; }
        public double ibaseise { get; set; }
        public double ise { get; set; }
        public string precio_medio { get; set; }
        public string tfactura { get; set; }
        public string testfact { get; set; }
        public string dapersoc { get; set; }
        public string cnifdnic { get; set; }
        public Char tindgcpy { get; set; }
        public Char tmodopta { get; set; }
        public string crefaext { get; set; }
        public string kperfact { get; set; }
        public string motivo_refacturacion { get; set; }
        public string submotivo { get; set; }
        public string tipo_comentario { get; set; }
        public string comentario_refact { get; set; }
        public int[] tconfact { get; set; }
        public double[] iconfact { get; set; }
        public double[] vpotcon1 { get; set; }
        public double[] vpotmax { get; set; }
        public double[] vconath { get; set; }
        public double[] vconrth { get; set; }
        public double vconrthp { get; set; }
        public double[] vexcere { get; set; }
        public double[] preciac { get; set; }
        public double[] precipo { get; set; }
        public double[] vexcepo { get; set; }
        public double vpotcall { get; set; }
        public double vpotcalv { get; set; }
        public double vpotcalp { get; set; }

        public FacturaInffact()
        {
            tconfact = new int[20];
            iconfact = new double[20];
        }
    }
}
