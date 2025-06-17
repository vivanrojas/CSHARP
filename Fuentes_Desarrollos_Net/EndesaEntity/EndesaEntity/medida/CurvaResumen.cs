using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class CurvaResumen
    {
        public string cups13 { get; set; }
        public string cups20 { get; set; }
        public string cups22 { get; set; }
        public string tarifa { get; set; }
        public DateTime fd { get; set; }
        public DateTime fh { get; set; }
        public int version { get; set; }
        public string estado { get; set; }
        public double activa { get; set; }
        public double reactiva { get; set; }        
        public double[] activa_periodo { get; set; }
        public double[] reactiva_periodo { get; set; }
        public double[] aci { get; set; }
        public double[] potencias_maximas { get; set; }
        public string origen { get; set; }
        public bool completa { get; set; }

        public int num_periodos { get; set; }

        public CurvaResumen()
        {
            activa_periodo = new double[7];
            reactiva_periodo = new double[7];
            aci = new double[7];
            potencias_maximas = new double[7];
        }

    }
}
