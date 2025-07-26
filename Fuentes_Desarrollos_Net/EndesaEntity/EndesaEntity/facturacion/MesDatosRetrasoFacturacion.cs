using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class MesDatosRetrasoFacturacion
    {
        public Int32 veces { get; set; }
        public Int32 mes { get; set; }
        public List<String> cupsree { get; set; }
        public Int32 dias { get; set; }
        public Int32 max { get; set; }

        public int year { get; set; }

        public MesDatosRetrasoFacturacion()
        {
            cupsree = new List<string>();
        }
    }
}
