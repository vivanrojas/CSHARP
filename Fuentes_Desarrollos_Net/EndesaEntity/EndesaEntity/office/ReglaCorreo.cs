using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.office
{
    public class ReglaCorreo
    {
        public string buzon { get; set; }
        public string carpeta { get; set; } // por si hay que leer dentro de una carpeta en concreto del buzon
        public string de { get; set; }
        public string para { get; set; }
        public string asunto { get; set; }
        public bool conAdjuntos { get; set; }
        public bool guardarAdjuntos { get; set; }
        public string rutaSalvadoAdjuntos { get; set; }
        public bool moverCorreo { get; set; }
        public string carpetaDestinoCorreo { get; set; }
        public bool ignorar_asunto { get; set; }
        public bool buzon_adicional { get; set; }
        public string carpetaDestinoAlternativo { get; set; }
    }
}
