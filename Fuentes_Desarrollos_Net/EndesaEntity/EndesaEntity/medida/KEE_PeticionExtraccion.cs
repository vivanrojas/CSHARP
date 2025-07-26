using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class KEE_PeticionExtraccion
    {
        public string cups { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public string origen { get; set; }
        public string tipo_curva { get; set; }
        public string motivo { get; set; }
        public string usuario { get; set; }
        public DateTime fecha_envio { get; set; }

    }
}
