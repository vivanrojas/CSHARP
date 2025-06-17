using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion
{
    public class Inventario_Detalle_Estados_Tabla : Audit
    {
        public string nif { get; set; }
        public string razon_social { get; set; }
        public string cups22 { get; set; }
        public int linea { get; set; }
        public string codigoproceso { get; set; }
        public string codigopaso { get; set; }
        public string codigosolicitud { get; set; }        
        public DateTime fechaactivacion { get; set; }
        public int estado { get; set; }
        public int subestado { get; set; }

    }
}
