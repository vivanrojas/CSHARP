using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion.cuadroDeMando
{
    public class Resumen
    {
        public int aniomes { get; set; }
        public string pais { get; set; }
        public string segmento { get; set; }
        public string grupo { get; set; }
        public int id_concepto { get; set; }
        public string concepto { get; set; }
        public int[] dias { get; set; }

        public Resumen()
        {
            dias = new int[32];
            for (int i = 1; i <= 31; i++)
                dias[i] = -1;
        }

    }
}
