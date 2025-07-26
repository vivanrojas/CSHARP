using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class PendienteMedida_B2B
    {
        public string comercializadora { get; set; }
        public string cups20 { get; set; }
        public string cups22 { get; set; }
        public string contrato_ps { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public int mes { get; set; }
        public string estado { get; set; }
        public string distribuidora { get; set; }
        public bool multipunto { get; set; }
        public string ritmo_facturacion { get; set; }
        public string tco_segm_back { get; set; }

        public DateTime fecha_informe { get; set; }
        public DateTime fec_act { get; set; }
        public int cod_carga { get; set; }
        public int id_pte_web { get; set; }

        public DateTime fecha_fin_KEE { get; set; }
        public DateTime fec_alta_kee { get; set; }

        public DateTime fecha_ControlCargas { get; set; }

        public DateTime fecha_modificacion { get; set; }

        public string cod_estado { get; set; }

        public string area { get; set; }
        public string incidencia { get; set; }

        public string cups { get; set; }

        public string titulo { get; set; }

        public string estado_incidencia { get; set; }

        public DateTime fecha_apertura { get; set; }

        public string prioridad_negocio { get; set; }

        public string e_s_estado { get; set; }

        public string Reincidente { get; set; }

        public string mes_pendiente { get; set; }

        public string id_crto_ext { get; set; }

        public DateTime fecha_baja_sap { get; set; }
    }
}

