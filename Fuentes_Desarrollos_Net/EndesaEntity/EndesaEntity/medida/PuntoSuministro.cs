using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class PuntoSuministro
    {
        public int id { get; set; }        
        public string cups13 { get; set; }
        public List<string> cups15 { get; set; }
        public List<string> cups22 { get; set; }
        public string cups20 { get; set; }
        public DateTime fd { get; set; }
        public DateTime fh { get; set; }

        public PuntoSuministro()
        {
            cups15 = new List<string>();
            cups22 = new List<string>();
        }
    }
}
