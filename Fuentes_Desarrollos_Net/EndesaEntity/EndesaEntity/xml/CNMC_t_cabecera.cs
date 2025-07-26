using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.xml
{
    public class CNMC_t_cabecera
    {
        public int id { get; set; }
        public string CodigoREEEmpresaEmisora { get; set; }
        public string CodigoREEEmpresaDestino { get; set; }
        public string CodigoDelProceso { get; set; }
        public string CodigoDePaso { get; set; }
        public string CodigoDeSolicitud { get; set; }
        public string SecuencialDeSolicitud { get; set; }
        public DateTime FechaSolicitud { get; set; }
        public string CUPS { get; set; }


    }
}
