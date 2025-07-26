using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Informe_contratos_gas_vencimiento : Table_atrgas_contratos_detalle
    {
        public string cliente { get; set; }
        public string distribuidora { get; set; }
        public string gestor { get; set; }
        public string responsable_territorial { get; set; }
        public string continua { get; set; }
    }
}
