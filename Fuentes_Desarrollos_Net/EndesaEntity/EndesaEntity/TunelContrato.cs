using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class TunelContrato: EndesaEntity.Audit
    {
        
        public string cliente { get; set; }
        public DateTime fecha_inicio_tunel { get; set; }
        public DateTime fecha_final_tunel { get; set; }
        public double consumo_referencia { get; set; }
        public double banda_inferior_pct { get; set; }
        public double banda_superior_pct { get; set; }
        public double banda_inferior_gwh { get; set; }
        public double banda_superior_gwh { get; set; }
        public double consumo_real_kwh { get; set; }
        public double consumo_real_gwh { get; set; }
        public bool aplica_tunel { get; set; }
        public bool formula_antigua { get; set; }
        public bool completo { get; set; }
        public int total_cups { get; set; }
        public int total_incompletos { get; set; }
        public string comentario { get; set; }
        public List<Tunel> lista { get; set; }

        public bool baja_tension { get; set; }
        public bool ltps { get; set; }
        public TunelContrato()
        {
            lista = new List<Tunel>();
        }
    }
}
