using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.sigame
{
    public class Consumos
    {
        public int id_pmedida { get; set; }
        public double consumo_tm { get; set; }
        public double consumo_ruta { get; set; }
        public double kwh_ruta { get; set; }
        public double kwh_tm { get; set; }
        public int mes_consumo { get; set; }
    }
}
