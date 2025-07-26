using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class InformeComparativaContratos
    {
        public string cups { get; set; }
        public string contrato { get; set; }
        public int version_cierre { get; set; }
        public int version_conversion { get; set; }
        public int tension { get; set; }

        public Dictionary<string, EndesaEntity.facturacion.InformeComparativaContratosDetalle> dic { get; set; }


        public InformeComparativaContratos()
        {
            dic = new Dictionary<string, InformeComparativaContratosDetalle>(); // Key = Variable
        }
    }
}
