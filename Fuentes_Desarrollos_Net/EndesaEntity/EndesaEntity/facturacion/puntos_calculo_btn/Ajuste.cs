using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EndesaEntity.facturacion.puntos_calculo_btn
{
    public class Ajuste
    {
        public string cupsree { get; set; }
        public string nif { get; set; }
        public DateTime finiajus { get; set; }
        public DateTime ffinajus { get; set; }
        public DateTime finisajus { get; set; }
        public DateTime fcreajus { get; set; }
        public DateTime fbajajus { get; set; }
        public int testajus { get; set; }
        public string[] parametros { get; set; }

        public Ajuste()
        {
            parametros = new string[4];
        }

    }
}
