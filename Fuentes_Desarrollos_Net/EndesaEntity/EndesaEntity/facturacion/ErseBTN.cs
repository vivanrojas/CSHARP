using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class ErseBTN
    {
        public string empresa { get; set; }
        public string ccounips { get; set; }
        public string dapersoc { get; set; }
        public string creferen { get; set; }
        public string secfactu { get; set; }
        public string cfactura { get; set; }
        public DateTime ffactura { get; set; }
        public DateTime ffactdes { get; set; }
        public DateTime ffacthas { get; set; }
        public double potencia { get; set; }
        public Int32 consumo { get; set; }
        public double ise { get; set; }
        public double ifactura { get; set; }
        public double tasa_dge { get; set; }
        public double base_iva_intermedio { get; set; }
        public double iva_intermedio { get; set; }
        public double base_iva_r { get; set; }
        public double iva_r { get; set; }
        public double base_iva_n { get; set; }
        public double iva_n { get; set; }
        public double cav { get; set; }
        public Int32 tfactura { get; set; }
        public string testfact { get; set; }
        public Int32 consumo_activa1 { get; set; }
        public Int32 consumo_activa2 { get; set; }
        public Int32 consumo_activa3 { get; set; }
        public Int32 consumo_activa4 { get; set; }
        public Int32 consumo_activa5 { get; set; }
        public Int32 consumo_activa6 { get; set; }
        public Int32 consumo_punta { get; set; }
        public Int32 consumo_llano { get; set; }
        public Int32 consumo_valle { get; set; }
        public Int32 consumo_supervalle { get; set; }
        public double iva { get; set; }
        public double iimpues3 { get; set; }

        public string cupsree { get; set; }
        public string cnifdnic { get; set; }

        public Dictionary<int, double> dic { get; set; }
        public ErseBTN()
        {
            dic = new Dictionary<int, double>();
        }
    }
}
