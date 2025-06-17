using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class InformeContratosComplementos
    {
        public string cups { get; set; }
        public string cups20 { get; set; }
        public string cups22 { get; set; }
        public string contrato { get; set; }
        public int version { get; set; }
        public int tension { get; set; }        

        public Dictionary<string, string> dic { get; set; }

        public InformeContratosComplementos()
        {
            dic = new Dictionary<string, string>();
        }

    }
}
