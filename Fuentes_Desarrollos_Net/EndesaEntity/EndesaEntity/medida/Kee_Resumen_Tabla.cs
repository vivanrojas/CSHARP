using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class Kee_Resumen_Tabla
    {
        public string cups20 { get; set; }
        public string cups22 { get; set; }
        public DateTime fd_extraccion_formulas { get; set; }
        public DateTime fh_extraccion_formulas { get; set; }
        public string fuente_extraccion_formulas { get; set; }
        public bool existe_resumen_sce_ml { get; set; }
        public DateTime fd_resumen_sce_ml { get; set; }
        public DateTime fh_resumen_sce_ml { get; set; }
        public double activa_resumen { get; set; }
        public double reactiva_resumen { get; set; }
        public bool existe_kee_ch { get; set; }
        public double activa_kre_ch { get; set; }
        public double reactiva_kre_ch { get; set; }
        public double dif_activa { get; set; }
        public double dif_reactiva { get; set; }

        public Kee_Resumen_Tabla()
        {
            existe_resumen_sce_ml = false;
            existe_kee_ch = false;
        }

    }

   
}
