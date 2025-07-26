using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion
{
    public class SOLATRMT_InformeATR_Carterizado: EndesaEntity.contratacion.SolATRMT
    {
        public string territorio { get; set; }
        public string responsable_territorial { get; set; }
        public string mail_responsable_territorial { get; set; }
    }
}
