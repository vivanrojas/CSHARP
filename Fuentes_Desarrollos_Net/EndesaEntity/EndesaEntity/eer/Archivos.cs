using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.eer
{
    public class Archivos
    {

        public int id_factura { get; set; }
        public string ruta_factura { get; set; }
        public string ruta_sstt { get; set; }
        public bool generado_pdf { get; set; }
    }
}
