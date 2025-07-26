using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion
{
    public class MotivosRechazo
    {
        public string cups { get; set; }
        public string clienteActualizado { get; set; }
        public long numSolAtr { get; set; }
        public DateTime fechaRechazoSol { get; set; }
        public string tipoSolicitud { get; set; }
        public string rechazoPdte { get; set; }
        public string motivos { get; set; }
        public string comentario { get; set; }
    }
}
