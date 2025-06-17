using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class Kee_Tabla
    {
        public string tiporeporte { get; set; }
        public DateTime fecha_recepcion { get; set; }
        public DateTime fecha_infome { get; set; }
        public string cups { get; set; }
        public DateTime fecha_inicio { get; set; }
        public DateTime fecha_fin { get; set; }
        public string sumatorio_ae_cch { get; set; }
        public string sumatorio_r1_cch { get; set; }
        public string sumatorio_ae_ch { get; set; }
        public string sumatorio_r1_ch { get; set; }
        public bool periodo_completo { get; set; }
        public DateTime fecha_min_cch { get; set; }
        public DateTime fecha_min_cc { get; set; }

        public DateTime fecha_max_cch { get; set; }
        public DateTime fecha_max_ch { get; set; }
        
        public Dictionary<DateTime, DateTime> dias_sin_medida_cch { get; set; }
        public Dictionary<DateTime, DateTime> dias_sin_medida_ch { get; set; }

        public string num_dias_cch { get; set; }
        public string num_dias_cc { get; set; }
        public string tipo_pm { get; set; }

        public List<int> num_huecos_cch { get; set; }
        public List<int> num_huecos_ch { get; set; }

        public Kee_Tabla()
        {
            periodo_completo = false;
            dias_sin_medida_cch = new Dictionary<DateTime, DateTime>();
            dias_sin_medida_ch = new Dictionary<DateTime, DateTime>();
            num_huecos_cch = new List<int>();
            num_huecos_ch = new List<int>();
        }

    }
}
