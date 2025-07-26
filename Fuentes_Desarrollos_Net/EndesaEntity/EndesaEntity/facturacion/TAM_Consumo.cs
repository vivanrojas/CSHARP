using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class TAM_Consumo
    {
        public string cupsree { get; set; }
        public DateTime ffactdes { get; set; }
        public DateTime ffacthas { get; set; }
        public int year { get; set; }
        public int month { get; set; }
        public double importe_factura { get; set; }
        public double consumo { get; set; } 

    }
}
