using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion.mes13
{
    public class Seguimiento
    {
        public int id { get; set; }
        public string segmento { get; set; }
        public string entidad { get; set; }
        public string linea_negocio { get; set; }
        public string empresa_titular { get; set; }
        public string nif { get; set; }
        public string nombre_cliente { get; set; }
        public string cups13 { get; set; }
        public string cups20 { get; set; }
        public string referencia { get; set; }        
        public int secuencial { get; set; }
        public string control { get; set; }
        public double estimacion_importe { get; set; }
        public double estimacion_base { get; set; }
        public double estimacion_impuestos { get; set; }
        public string nif_referencia {get;set;}
        public DateTime diaf { get; set; }
        public DateTime fecha_vto_estimada { get; set; }
        public DateTime diav { get; set; }
        public int seguimiento { get; set; }
        public long creferen { get; set; }
        public int secfactu { get; set; }
        public string cfactura { get; set; }
        public DateTime ffactura { get; set; }
        public DateTime ffactdes { get; set; }
        public DateTime ffacthas { get; set; }
        public DateTime flimpago { get; set; }
        public DateTime fptacobr { get; set; }
        public double ifactura { get; set; }
        public double impuestos { get; set; }

    }
}
