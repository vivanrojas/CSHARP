using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion.cuadroDeMando
{
    public class InformeGas
    {
        public int id_PS { get; set; }
        public string cif { get; set; }
        public string nombre_punto_suminsitro { get; set; }
        public string cups { get; set; }
        public string grupo { get; set; }
        public DateTime fecha_inicio { get; set; }
        public DateTime fecha_fin { get; set; }

        public int ultimo_mes_facturado { get; set; }
        public DateTime medida { get; set; }

        public DateTime fecha_emision_sigame { get; set; }
        public DateTime fecha_informe { get; set; }
        public DateTime ffactura { get; set; }
        public string pendiente { get; set; }
        public string comentario { get; set; }
        public double promedio_facturacion { get; set; }

        public bool es_cisterna { get; set; }

        public string pais { get; set; }

    }
}
