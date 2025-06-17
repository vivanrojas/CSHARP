using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Medida
    {
        public string cups13 { get; set; }
        public string cups20 { get; set; }
        public DateTime fromdate { get; set; }
        public DateTime todate { get; set; }
        public string estado { get; set; }
        public double activa { get; set; }
        public double reactiva { get; set; }
        public string fuente { get; set; }
        public int dias { get; set; }

        public List<medida.CurvaResumen> lista_curva_resumen { get; set; }

        public Medida()
        {
            lista_curva_resumen = new List<medida.CurvaResumen>();
        }
    }
}
