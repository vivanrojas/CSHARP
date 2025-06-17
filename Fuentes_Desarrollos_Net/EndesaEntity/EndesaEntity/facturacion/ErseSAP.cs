using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class ErseSAP
    {
        public string de_seg_merc_por { get; set; }
        public string cd_estado_fact {  get; set; }
        public DateTime fh_ult_ejec { get; set; }
        public string cif { get; set; }
        public string cfactura { get; set; }
        public DateTime ffactura { get; set; }
        public DateTime ffactdes { get; set; }
        public DateTime ffacthas { get; set; }
        public double ifactura { get; set; }
        public double impuesto1 { get; set; }
        public double consumo { get; set; }
        public double ise { get; set; }
        public double base_iva_normal { get; set; }
        public double iva_normal { get; set; }
        public double base_iva_reducido { get; set; }
        public double iva_reducido { get; set; }
        public double cav { get; set; }
        public double importe_redes { get; set; }
        public double consumo_activa1 { get; set; }
        public double consumo_activa2 { get; set; }
        public double consumo_activa3 { get; set; }
        public double consumo_activa4 { get; set; }
        public string cupsree { get; set; }
        public double potencia { get; set; }
    }
}
