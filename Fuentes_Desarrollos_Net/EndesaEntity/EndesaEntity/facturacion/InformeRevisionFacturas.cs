using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class InformeRevisionFacturas
    {
        public int empresa { get; set; }
        public string cif { get; set; }
        public string cliente { get; set; }
        public string cups22 { get; set; }
        public string tarifa { get; set; }
        public string provincia { get; set; }
        public double cod_sol_atr {get;set;}
        public string estado_sol_atr { get; set; }
        public string tipo_sol_atr { get; set; }
        public DateTime fecha_activacion { get; set; }
        public bool agora { get; set; }
        public bool agrupada { get; set; }
        public int tipo_gestion_atr { get; set; }
        public int tpuntmed { get; set; }
        public string tipo_contrato { get; set; }
        public string complemento { get; set; }
        public DateTime fecha_anexion { get; set; }


    }
}
