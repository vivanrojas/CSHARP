using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.punto_suministro
{
    public class PuntoSuministro
    {
        public string cups20 { get; set; }
        public EndesaEntity.Direccion direccion { get; set; }

        // public Dictionary<string, string> dic_puntos_medida_principales { get; set; }
        public List<EndesaEntity.medida.PuntoMedida> lista_puntos_medida_principales { get; set; }
        public EndesaEntity.punto_suministro.Tarifa tarifa { get; set; }
        public int[] potecias_contratadas { get; set; }
        public bool medida_en_baja { get; set; }
        public DateTime fecha_inicio { get; set; }
        public DateTime fecha_fin { get; set; }
        public double porcetaje_perdidas { get; set; }
        public double potencia_trafo { get; set; }
        public List<double> lista_cc { get; set; }
        public bool sstt { get; set; }
        public double importe_sstt { get; set; }
        public int cuotas { get; set; }
        public int cuotas_pdtes { get; set; }
        public double alquiler { get; set; }

        public double servicio_gestion_preferente { get; set; }
        public double dto_te { get; set; }

        public bool exencion_ise { get; set; }
        public double porcentaje_exencion { get; set; }
        public int tipo_punto_medida { get; set; }
        public EndesaEntity.PreciosEnergia precios_energia { get; set; }

        public Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga> dic_cc;

        // public EndesaEntity.medida.CurvaVector curvaVector { get; set; }
               
        public PuntoSuministro()
        {
            potecias_contratadas = new int[7];
            direccion = new Direccion();
            // dic_puntos_medida_principales = new Dictionary<string, string>();
            lista_puntos_medida_principales = new List<medida.PuntoMedida>();
            lista_cc = new List<double>();
            dic_cc = new Dictionary<DateTime, medida.CurvaDeCarga>();

        }
    }
}
