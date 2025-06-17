using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.gas
{
    public class XML_GAS
    {
        public string codigo_envio { get; set; }
        public string empresa_emisora_del_paso { get; set; }
        public string empresa_receptora_del_paso { get; set; }
        public DateTime fecha_hora_comunicacion { get; set; }
        public string codigo_proceso { get; set; }
        public string codigo_tipo_mensaje { get; set; }
        public string detalle_solicitud_nuevo_contrato_producto { get; set; }
        public long referencia_solicitud_comercializadora { get; set; }
        public DateTime fecha_hora_solicitud { get; set; }
        public DateTime tipo_titular { get; set; }
        public string nacionalidad { get; set; }
        public string tipo_documento_idenficacion { get; set; }
        public string documento { get; set; }
        public string cups { get; set; }
        public string modelo_fecha_efecto { get; set; }
        public DateTime fecha_efecto_solicitada { get; set; }
        public string tipo_solicitud { get; set; }
        public string tipo_producto { get; set; }
        public string tipo_peaje_producto { get; set; }
        public double qd { get; set; }

    }
}
