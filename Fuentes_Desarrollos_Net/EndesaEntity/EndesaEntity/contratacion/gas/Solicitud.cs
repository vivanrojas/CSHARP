using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.gas
{
    public class Solicitud
    {

        public string razon_social { get; set; }
        public string nif { get; set; }
        public string calle { get; set; }
        public string numero { get; set; }
        public string municipio { get; set; }
        public string provincia { get; set; }
        public string codigo_postal { get; set; }
        public string cups { get; set; }
        public string grupo_tarifario { get; set; }
        public string mail_remitente { get; set; }
        public DateTime fecha_mail { get; set; }
        public string asunto_mail { get; set; }
        public bool hayError { get; set; }
        public string descripcion_error { get; set; }
        public string mail_distribuidora { get; set; }
        public  string distribuidora_tramitacion { get; set; }
        public string descripcion { get; set; }
        

        public List<SolicitudDetalle> detalle { get; set; }

        public Solicitud()
        {
            hayError = false;
            detalle = new List<SolicitudDetalle>();
        }
    }
}
