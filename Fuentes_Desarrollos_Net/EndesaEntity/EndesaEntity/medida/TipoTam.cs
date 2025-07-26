using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class TipoTam
    {
        public int cemptitu { get; set; }
        public string cnifdnic { get; set; }
        public string dapersoc { get; set; }
        public string cups13 { get; set; }
        public string cups20 { get; set; }
        public double tam { get; set; }
        public int numfacturasCalculo { get; set; }
        public double importeMaximo { get; set; }
        public int diasFacturacionMedia { get; set; }
        public int diasFacturacionMaxima { get; set; }
        public DateTime fechaHastaMaxima { get; set; }
        public int rankingTAM { get; set; }
        public int rankingImporteMaximo { get; set; }
        public bool esPS_AT { get; set; }

        public TipoTam()
        {
            esPS_AT = false;
            tam = 0;
            numfacturasCalculo = 0;
            importeMaximo = 0;
            diasFacturacionMedia = 0;
            diasFacturacionMaxima = 0;
            rankingTAM = 0;
            rankingImporteMaximo = 0;
        }
    }
}
