using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion.gestionATRGas
{
    public class ContratoGasDetalle : EndesaEntity.Table_atrgas_contratos_detalle
    {
        public string estado { get; set; }
        public int ultimo_mes_facturado { get; set; }
        public string estadoContrato { get; set; }
        public string vatnum { get; set; }
        public string customer_name { get; set; }
    }
}
