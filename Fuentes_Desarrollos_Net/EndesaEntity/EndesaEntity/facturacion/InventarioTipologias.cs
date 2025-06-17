using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EndesaEntity.facturacion
{
    public class InventarioTipologias
    {
        public string cups13 { get; set; }
        public string cups22 { get; set; }        
        public DateTime fecha_alta { get; set; }
        public DateTime fecha_baja { get; set; }
        public string cif { get; set; }
        public string razon_social { get; set; }
        public bool convertido { get; set; }
        public string empresa { get; set; }
        public string tarifa { get; set; }
        public string tipo_tarifa { get; set; }
        public string tipo_punto_suministro { get; set; }
        public string provincia { get; set; }
        public string territorio { get; set; }
        public string contrato_ext { get; set; }
        public bool agrupada { get; set; }
        public bool age { get; set; }
        public bool agora { get; set; }
        public bool passthough { get; set; }
        public bool passpool_horario { get; set; }
        public string passpool { get; set; }
        public bool passpool_periodo { get; set; }
        public bool passpool_subasta { get; set; }
        public string tipo_contrato { get; set; }
        public bool gestion_propia_atr { get; set; }
        public double tam { get; set; }
        public string estado_contrato { get; set; }
        public int version { get; set; }
        public DateTime fecha_puesta_servicio { get; set; }
        public bool revendedores { get; set; }
        public bool ultimo_convertido { get; set; }
        public bool medido_en_baja { get; set; }
        public bool multipunto { get; set; }
        public double kaveas { get; set; }
        public string ciclo_p01_mt { get; set; }
        public string tipo_exencion { get; set; }
        public string iva_intermedio_btn { get; set; }
        public string iva_reducido { get; set; }
        public Dictionary<string, EndesaEntity.contratacion.ComplementosContrato> dic_complementos { get; set; }
        public string periodo_pendiente { get; set; }
        public string estado_pendiente { get; set; }
        public string subestado_pendiente { get; set; }

        public InventarioTipologias()
        {
            dic_complementos = new Dictionary<string, EndesaEntity.contratacion.ComplementosContrato>();
        }


    }
}
