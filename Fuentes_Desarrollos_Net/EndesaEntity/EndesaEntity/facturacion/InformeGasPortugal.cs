using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class InformeGasPortugal
    {
        public int id_ps { get; set; }
        public string cups20 { get; set; }
        public string dapersoc { get; set; }
        public string nif { get; set; }
        public DateTime ffactdes { get; set; }
        public DateTime ffacthas { get; set; }
        public string cfactura { get; set; }
        public DateTime ffactura { get; set; }
        public double importe_neto { get; set; }
        public double importe_iva { get; set; }
        public double importe_iva_reducido { get; set; }
        public double importe_bruto { get; set; }
        public double qfacturada { get; set; }
        public double consumo { get; set; }
        public double consumo_m3 { get; set; }
        public double tos_importe { get; set; }
        public double ieh_importe { get; set; }
        public double consumo_kwh { get; set; }
        public double dif { get; set; }
        public double import_acceso_redes { get; set; }
        public string moneda { get; set; }
        public double importe_dto { get; set; }
        public double kw_dto { get; set; }

        public bool lg_pf { get; set; }
        public bool lg_pool { get; set; }
        public bool lg_brent { get; set; }

    }
}
