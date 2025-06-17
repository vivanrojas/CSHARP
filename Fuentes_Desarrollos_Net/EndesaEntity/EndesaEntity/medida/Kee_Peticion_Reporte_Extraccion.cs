using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class Kee_Peticion_Reporte_Extraccion
    {
        public string cups20 { get; set; }
        public string fuente { get; set; }
        public string tipo { get; set; }
        public List<Kee_Reporte_Extraccion> lista { get; set; }

        public Kee_Peticion_Reporte_Extraccion()
        {
            lista = new List<Kee_Reporte_Extraccion>();
        }
    }
}
