using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.eer
{
    public class FacturaDetalle
    {
        public int id_factura { get; set; }
        public string numero_factura { get; set; }
        public int linea_factura { get; set; }
        public string concepto { get; set; }
        public string producto { get; set; }
        public string calculo { get; set; }
        public List<string> lista_detalle_calculo { get; set; }
        public string descripcion { get; set; }
        public double cantidad { get; set; }
        public string unidad_cantidad { get; set; }
        public double precio { get; set; }
        public string unidad_precio { get; set; }
        public double total { get; set; }

        public FacturaDetalle()
        {
            lista_detalle_calculo = new List<string>();
        }



    }
}
