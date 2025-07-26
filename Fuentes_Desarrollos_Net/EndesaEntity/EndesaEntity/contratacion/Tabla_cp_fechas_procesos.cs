using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion
{
    public class Tabla_cp_fechas_procesos
    {
        public string archivo { get; set; }
        public string proceso { get; set; }
        public string extractor { get; set; }
        public DateTime fecha   { get; set; }
        public long kb { get; set; }
        public string md5 { get; set; }


    }
}
