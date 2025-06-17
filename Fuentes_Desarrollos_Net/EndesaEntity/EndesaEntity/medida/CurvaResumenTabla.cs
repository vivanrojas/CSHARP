using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class CurvaResumenTabla    
    {
        public string cpuntmed { get; set; }
        public string fuente { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public string cups13 { get; set; } // cups13
        public string cups20 { get; set; }  // cups20
        public string cups22 { get; set; } // cups22
        public int dias { get; set; } // dias
        public string estado { get; set; } // R, F, D, A
        public double activa { get; set; } // ACTIVA
        public double reactiva { get; set; } // REACTIVA      
        public int version { get; set; }
        public bool existe_curva { get; set; }

        public List<CurvaCuartoHorariaTabla> curvasCuartoHorarias {get;set;}

        public CurvaResumenTabla()
        {
            existe_curva = false;
            curvasCuartoHorarias = new List<CurvaCuartoHorariaTabla>();
        }
    }
}
