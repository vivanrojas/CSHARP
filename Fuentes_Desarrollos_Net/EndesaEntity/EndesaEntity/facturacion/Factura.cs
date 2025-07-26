using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class Factura
    {
        public string entorno { get; set; }
        public int id_pto_suministro { get; set; } // para gas
        public DateTime fecha_expedicion_factura { get; set; } // para gas
        public bool origen_sce { get; set; } // para gas
        public int cemptitu { get; set; }
        public string descripcion_empresa { get; set; }
        public string cnifdnic { get; set; }
        public string dapersoc { get; set; }
        public string ccounips { get; set; }
        public string cupsree { get; set; }
        public string cfactura { get; set; }
        public DateTime ffactura { get; set; }
        public DateTime ffactdes { get; set; }
        public DateTime ffacthas { get; set; }
        public double ifactura { get; set; }
        public double ise { get; set; }
        public double ibaseise { get; set; }
        public double iva { get; set; }
        public double iimpues2 { get; set; }
        public double iimpues3 { get; set; }
        public long creferen { get; set; }
        public int secfactu { get; set; }
        public int tfactura { get; set; }
        public string tfactura_descripcion { get; set; }
        public string tfactura_desc { get; set; }
        public string testfact { get; set; }
        public string tipo_negocio { get; set; }
        public int vcuovafa { get; set; }
        public int tconfac { get; set; }
        public string tconfac_descripcion_corta { get; set; }
        public string tconfac_descripcion_larga { get; set; }
        public double iconfac { get; set; }
        public int ultimo_mes_facturado { get; set; }
        public bool existe { get; set; }
        public bool facturado { get; set; }

        public string comentario { get; set; }

        public string cfacagp { get; set; }

        public double cnpr { get; set; }


        public List<Factura_Detalle> lista_conceptos {get; set;}
        public string cfactrec { get; set; }
        public Factura()
        {            
            lista_conceptos = new List<Factura_Detalle>();
        }
    }
}
