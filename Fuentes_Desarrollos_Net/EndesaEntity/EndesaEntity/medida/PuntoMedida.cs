using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class PuntoMedida
    {
        public string cups22 { get; set; }
        public DateTime fechaVigor { get; set; }
        public DateTime fechaAlta { get; set; }
        public DateTime fechaBaja { get; set; }
        public int tipoPuntoMedida { get; set; }
        public int modoLectura { get; set; }
        public string funcion { get; set; }
        public int direccion_enlace { get; set; }
        public int direccion_punto_medida { get; set; }
        public int telefono_telemedida { get; set; }
        public int clave { get; set; }
        public int estadoTelefono { get; set; }
        public bool marcaMedidaConPerdidas { get; set; }



    }
}
