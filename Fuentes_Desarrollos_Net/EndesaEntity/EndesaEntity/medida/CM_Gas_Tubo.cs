using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class CM_Gas_Tubo
    {
        public string cups { get; set; }
        public string pais { get; set; }
        public string nombre_cliente { get; set; }
        public string distribuidora { get; set; }
        public bool telemedida { get; set; }
        public bool top { get; set; }
        public int primer_mes_pdte { get; set; }
        public int mes { get; set; }
        public DateTime fecha_alta { get; set; }
        public string area_pdte { get; set; }
        public string motivo_pdte { get; set; }
        public double tam { get; set; }
    }
}
