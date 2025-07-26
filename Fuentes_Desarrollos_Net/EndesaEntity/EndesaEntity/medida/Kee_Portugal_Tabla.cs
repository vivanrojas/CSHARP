using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class Kee_Portugal_Tabla
    {
        public EndesaEntity.medida.Kee_TipoReporte tiporeporte { get; set; }
        public DateTime fecha_recepcion { get; set; }
        public DateTime fecha_infome { get; set; }
        public string cups { get; set; }
        public DateTime fecha_inicio { get; set; }
        public DateTime fecha_fin { get; set; }
        public bool definitivo { get; set; }
        public string sumario_ae { get; set; }
        public string sumatorio_ri { get; set; }
        public string sumatorio_rc { get; set; }        
        public bool periodo_completo { get; set; }
        public DateTime fecha_max_cch { get; set; }
        

        public Dictionary<DateTime, DateTime> dias_incompletos { get; set; }
        

        public Kee_Portugal_Tabla()
        {
            definitivo = false;
            periodo_completo = false;
            dias_incompletos = new Dictionary<DateTime, DateTime>();
            
        }
    }
}
