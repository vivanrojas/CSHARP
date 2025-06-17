using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class InventarioFacturacion
    {
        public string error { get; set; }
        public bool actualizado { get; set; }
        public string nif { get; set; }
        public string  cliente { get; set; }
        public string carpeta_cliente { get; set; }
        public string cups13 { get; set; }
        public string cpe { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public string ruta_plantilla { get; set; }
        public string ltp { get; set; }
        public string estado { get; set; }
        



    }
}
