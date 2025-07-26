using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class PNT_XML_Detail
    {
        public DateTime fechaInicioValoracion { get; set; }
        public DateTime fechaFinValoracion { get; set; }
        public int[] potenciaAFacturar { get; set; }
        public int[] energiaAFacturar { get; set; }
        public int[] excesoPotAFacturar { get; set; }
        public int[] reactivaAFacturar { get; set; }

        public PNT_XML_Detail()
        {
            potenciaAFacturar = new int[6];
            energiaAFacturar = new int[6];
            excesoPotAFacturar = new int[6];
            reactivaAFacturar = new int[6];

        }
    }
}
