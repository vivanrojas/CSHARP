using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion.puntos_calculo_btn
{
    public class puntos_calculo
    {
        public string cpe { get; set; }
        public string contrato { get; set; }
        public double pct_aplicacion { get; set; }

        public DateTime f_desde { get; set; }
        public DateTime f_hasta { get; set; }
        public int consumo { get; set; }
        public string perfil { get; set; }
        public string calendario { get; set; }
        public string tarifa { get; set; }
        public double kWh_afectados { get; set; }
        public DateTime f_generacion { get; set; }
        public string tipo { get; set; }
        public bool existe { get; set; }
        

        public puntos_calculo()
        {
            existe = false;
        }

    }
}
