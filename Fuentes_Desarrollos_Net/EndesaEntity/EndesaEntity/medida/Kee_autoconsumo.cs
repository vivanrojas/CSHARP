using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class Kee_autoconsumo
    {
        public string cups22 { get; set; }
        public int tipo_autoconsumo { get; set; }
        public string tipo_curva { get; set; }
        public DateTime fecha_hora_lectura { get; set; }
        public int consumo_ae { get; set; }
        public int consumo_as { get; set; }
        public string archivo_origen { get; set; }

    }
}
