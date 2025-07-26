using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion
{
    public class Exenciones: EndesaEntity.Audit
    {
        public string tipo_ajuste { get; set; }
        public string tipo_identificador_fiscal { get; set; }
        public string identificador_fiscal { get; set; }
        public string cups13 { get; set; }
        public string cups22 { get; set; }
        public string razon_social { get; set; }
        public string nombre { get; set; }
        public DateTime fecha_inicio_vigencia { get; set; }
        public DateTime fecha_fin_vigencia { get; set; }
        public string num_dias_hasta_caducar { get; set; }
        public int porcentaje_exento { get; set; }
        public double porcentaje_aplicacion { get; set; }
        public string codigo_tarjeta { get; set; }
        public string articulo_cie { get; set; }

    }
}
