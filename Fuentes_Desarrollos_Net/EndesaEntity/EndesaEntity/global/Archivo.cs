using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.global
{
    public class Archivo
    {
        public string nombre_archivo { get; set; }
        public string ruta_completa { get; set; }
        public double size { get; set; }
        public DateTime fechaUltimoAcceso { get; set; }
        public DateTime fechaCreaccion { get; set; }
        public DateTime fechaModificacion { get; set; }
    }
}
