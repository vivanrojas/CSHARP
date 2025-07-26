using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Solicitud
    {
        public int id { get; set; }
        public string mail { get; set; }
        public DateTime fechahora_mail { get; set; }
        public string desc_error { get; set; }
    }
}
