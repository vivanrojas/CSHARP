using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.eer
{
    public class Factura
    {

        public int id_factura { get; set; }
        public string nif { get; set; }
        public string nombre_cliente { get; set; }
        public string direccion_facturacion { get; set; }
        public string direccion_envio { get; set; }
        public string direccion_suministro { get; set; }

        public string codigo_factura { get; set; }
        public DateTime fecha_factura { get; set; }
        public DateTime fecha_consumo_desde { get; set; }
        public DateTime fecha_consumo_hasta { get; set; }

        public int dias_facturados { get; set; }
        public string cupsree { get; set; }
        public string tarifa { get; set; }
        public double consumo_activa { get; set; }
        public double consumo_reactiva { get; set; }


        // Importes

        public double base_ise_reducido { get; set; }
        public double iva { get; set; }
        public double impuesto_electricidad { get; set; }
        public double impuesto_electricidad_reducido { get; set; }
        public double base_ise { get; set; }
        public double base_imponible { get; set; }
        public double base_imponible_ie { get; set; }
        public double total_factura { get; set; }

        // Energia
        public double[] consumos_periodos_activa { get; set; }
        public double[] consumos_periodos_reactiva { get; set; }
        public double[] potencias { get; set; }

        public double termino_energia { get; set; }
        public double descuento_energia { get; set; }
        public double facturacion_potencia { get; set; }
        public double recargo_excesos { get; set; }

        public double recargo_excesos_reactiva { get; set; }


        public double suplemento_territorial { get; set; }
        public double alquiler { get; set; }

        public List<FacturaDetalle> lista_factura_detalle { get; set; }

        public Factura()
        {
            consumos_periodos_activa = new double[7];

            consumos_periodos_reactiva = new double[7];

            potencias = new double[7];

            lista_factura_detalle = new List<FacturaDetalle>();
        }



    }
}
