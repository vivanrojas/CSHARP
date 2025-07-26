using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.global
{
    public class MSAccess: Archivo
    {
        public int tablas_locales { get; set; }
        public int tablas_externas { get; set; }
        public int total_modulos { get; set; }
    }
}
