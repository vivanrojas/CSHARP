using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Inventario
    {
        public string nif { get; set; }
        public string nombre_cliente { get; set; }
        public Dictionary<string, EndesaEntity.punto_suministro.PuntoSuministro> dic_puntos_suministro { get; set; }
        public DateTime fecha_inicio { get; set; }
        public DateTime fecha_fin { get; set; }
        
        public DateTime fecha_alta { get; set; }
        public DateTime fecha_baja { get; set; }               
        
        public DateTime fecha_factura { get; set; }
        
        public string tipo_contrato { get; set; }
        public string tipo_autoconsumo { get; set; }
        public string cnae { get; set; }
               
        public string direccion_facturacion { get; set; }
        public string direccion_envio { get; set; }

        public Inventario()
        {
            dic_puntos_suministro = new Dictionary<string, punto_suministro.PuntoSuministro>();
        }
       

    }
}
