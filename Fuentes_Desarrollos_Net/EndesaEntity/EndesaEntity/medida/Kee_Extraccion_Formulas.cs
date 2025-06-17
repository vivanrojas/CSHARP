using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class Kee_Extraccion_Formulas
    {
        public string cups20 { get; set; }
        public string cups22 { get; set; }

        public DateTime fecha_sol_desde { get; set; }
        public DateTime fecha_sol_hasta { get; set; }

        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public string tipo { get; set; }
        public string fuente { get; set; }
        public string extraccion { get; set; }
        public DateTime fecha_mod_archivo { get; set; }
        public string usuario { get; set; }
        public List<Kee_Reporte_Extraccion> lista { get; set; }
                

        public Kee_Extraccion_Formulas()
        {
            lista = new List<Kee_Reporte_Extraccion>();
        
        }



    }
}
