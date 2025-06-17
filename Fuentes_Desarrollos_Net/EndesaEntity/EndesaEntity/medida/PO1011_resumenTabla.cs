using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class PO1011_resumenTabla
    {
        public string cups13 { get; set; }
        public string cups15 { get; set; }
        public string cups20 { get; set; }
        public string cups22 { get; set; }
        public DateTime fd_vigencia_pm { get; set; }
        public DateTime fh_vigencia_pm { get; set; }
        public DateTime ffactdes_resumen { get; set; }
        public DateTime ffacthas_resumen { get; set; }
        public string testrcur { get; set; }
        public double activa_resumen { get; set; }
        public double reactiva_resumen { get; set; }

       

        public double activa_p1 { get; set; }
        public double reactiva_p1 { get; set; }
        public double activa_f1 { get; set; }
        public double reactiva_f1 { get; set; }
        public int periodos_p1 { get; set; }

        // P2
        public double activa_p2 { get; set; }
        public double reactiva_p2 { get; set; }

        public int periodos_p2 { get; set; }
        public bool completo_p2 { get; set; }
        public bool encontrado_p2 { get; set; }


        public int periodos_f1 { get; set; }
        public bool completo_p1 { get; set; }
        public bool completo_f1 { get; set; }
        public string origen { get; set; }
        public bool encontrado_p1 { get; set; }
        public bool encontrado_f1 { get; set; }
        public PO1011_resumenTabla()
        {
            encontrado_p1 = false;
            encontrado_f1 = false;
            completo_p1 = false;
            completo_f1 = false;
        }
    }
}


