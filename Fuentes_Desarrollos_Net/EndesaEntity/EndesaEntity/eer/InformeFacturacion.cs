using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.eer
{
    public class InformeFacturacion
    {
        public string cups { get; set; }
        public string razon_social { get; set; }
        public string cif { get; set; }
        public string domicilio { get; set; }
        public string municipio { get; set; }
        public double[] potencia { get; set; }
        public string tarifa_acceso { get; set; }
        public string referencia_contratato { get; set; }
        public string codigo_factura { get; set; }
        public DateTime fecha_consumo_desde { get; set; }
        public DateTime fecha_consumo_hasta { get; set; }
        public DateTime fecha_emision { get; set; }
        public double[] consumo_activa { get; set; }
        public double[] consumo_reactiva { get; set; }
        public double[] potencia_demandada { get; set; }
        public double coste_de_potencia { get; set; }
        public double excesos_de_potencia { get; set; }
        public double importe_energia { get; set; }
        public double excesos_de_reactiva { get; set; }
        public double impuesto_electrico { get; set; }
        public double impuesto_electrico_reducido { get; set; }
        public double alquiler_equipo { get; set; }
        public double fianza { get; set; }
        public double derechos_contratacion { get; set; }
        public double importe_sin_iva { get; set; }
        public double iva { get; set; }
        public double total_factura { get; set; }

        public InformeFacturacion()
        {
            potencia = new double[7];
            consumo_activa = new double[7];
            consumo_reactiva = new double[7];
            potencia_demandada = new double[7];
        }
    }
}
