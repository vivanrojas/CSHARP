using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion.mes13
{
    public  class FacturaMes13_Seguimiento
    {
        public enum linea_negocio
        {
            Luz,
            Gas
        }

        public String ccounips { get; set; }
        public String cupsree { get; set; }
        public string cemptitu { get; set; }
        public long creferen { get; set; }
        public Int32 secfactu { get; set; }
        public String creffact { get; set; }
        public String cfactura { get; set; }
        public DateTime ffactura { get; set; }
        public DateTime ffactdes { get; set; }
        public DateTime ffacthas { get; set; }
        public Double vcuovafa { get; set; }
        public String venereac { get; set; }
        public Decimal vcuofifa { get; set; }
        public Decimal ifactura { get; set; }
        public Decimal iva { get; set; }
        public Decimal iimpues2 { get; set; }
        public Decimal iimpues3 { get; set; }
        public Decimal ibaseise { get; set; }
        public Decimal ise { get; set; }
        public String precio_medio { get; set; }
        public string tfactura { get; set; }
        public String testfact { get; set; }
        public String dapersoc { get; set; }
        public String cnifdnic { get; set; }
        public Char tindgcpy { get; set; }
        public Char tmodopta { get; set; }
        public String crefaext { get; set; }
        public String kperfact { get; set; }
        public String motivo_refacturacion { get; set; }
        public String submotivo { get; set; }
        public String tipo_comentario { get; set; }
        public string comentario_refact { get; set; }
        public int[] tconfact { get; set; }
        public Decimal[] iconfact { get; set; }
        public decimal[] vpotcon1 { get; set; }
        public decimal[] vpotmax { get; set; }
        public decimal[] vconath { get; set; }
        public decimal[] vconrth { get; set; }
        public decimal vconrthp { get; set; }
        public decimal[] vexcere { get; set; }
        public decimal[] preciac { get; set; }
        public decimal[] precipo { get; set; }
        public decimal[] vexcepo { get; set; }
        public decimal vpotcall { get; set; }
        public decimal vpotcalv { get; set; }
        public decimal vpotcalp { get; set; }

        public FacturaMes13_Seguimiento()
        {
            tconfact = new int[20];
            iconfact = new decimal[20];
        }
    }
}
