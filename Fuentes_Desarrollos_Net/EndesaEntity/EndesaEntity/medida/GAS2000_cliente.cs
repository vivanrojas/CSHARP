using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class GAS2000_cliente
    {
        public string cliente { get; set; }
        public string otro_nombre { get; set; }
        public string cif { get; set; }
        public int id_pto_suministro { get; set; }
        public string cups { get; set; }
        public string gestor { get; set; }
        public string poblacion { get; set; }
        public string provincia {get;set;}
        public string manda { get; set; }
        public string persona { get; set; }
        public string tipo_envio { get; set; }
        public string factura { get; set; }
        public string distribuidora { get; set; }
        public bool cisternas { get; set; }
        public bool consumoscliente { get; set; }
        public int qcontratado1 { get; set; }
        public DateTime fecha_alta { get; set; }
        public DateTime fecha_baja { get; set; }
        public int um { get; set; }
        public bool telemedida { get; set; }
        public bool top { get; set; }
        public string motivo_pendiente { get; set; }

    }
}
