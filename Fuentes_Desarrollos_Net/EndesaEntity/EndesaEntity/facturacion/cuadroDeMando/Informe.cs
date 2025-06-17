using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion.cuadroDeMando
{
    public class Informe
    {
        public string cd_pais { get; set; }
        public string empresa { get; set; }
        public int id_ps { get; set; }
        public string nif { get; set; }
        public string cliente { get; set; }
        public string grupo { get; set; }
        public string cups13 { get; set; }
        public string cups20 { get; set; }
        public string estado_contrato { get; set; }
        public string provincia { get; set; }
        public int mes { get; set; }
        public string estado_factura { get; set; }
        public string estado_LTP { get; set; }
        public DateTime fecha_desde_TLP { get; set; }
        public string tipo { get; set; }
        public string tipo_CM { get; set; }
        public double tam { get; set; }
        public double consumo_medio { get; set; }
        public string top { get; set; }
        //public string scoring { get; set; }
        public bool riesgo { get; set; }
        public string gestor { get; set; }
        public string posicion_gestor { get; set; }
        public string[] anomalia { get; set; }
        public string alarma { get; set; }
        public DateTime fecha_prevista_baja { get; set; }
        public DateTime fecha_emision_sigame { get; set; }
        public DateTime fecha_inicio_contrato { get; set; }
        public DateTime fecha_fin_contrato { get; set; }
        public DateTime fecha_informe { get; set; }
        public DateTime fecha_factura_sce { get; set; }
        public DateTime medida { get; set; }
        public string contratoPS { get; set; }
        public string origen_sistemas { get; set; }
        public DateTime fecha_baja_contrato { get; set; }






        public Informe()
        {
            anomalia = new string[5];
        }

    }
}
