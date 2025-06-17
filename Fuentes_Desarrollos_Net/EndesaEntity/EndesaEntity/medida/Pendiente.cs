using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class Pendiente: Audit
    {
        public string cod_empresaTitular { get; set; }
        public string empresaTitular { get; set; }
        public string segmento { get; set; }
        public string cups15 { get; set; }
        public string cups13 { get; set; }         
        public string cups20 { get; set; }
        public string codContrato { get; set; }
        public int aaaammPdte { get; set; }
        public string distribuidora { get; set; }
        public string cod_estado { get; set; }
        public string cod_subestado { get; set; }
        public string estado { get; set; } 
        public string subsEstado { get; set; }
        public string descripcion_estado { get; set; }
        public string descripcion_subestado { get; set; }
        public bool multimedida { get; set; }
        public DateTime fh_desde { get; set; }
        public DateTime fecha_informe { get; set; }
        public int num_cups { get; set; }
        public bool existe { get; set; }
        public double tam { get; set; }
        public bool agora { get; set; }
        public string area_responsable { get; set; }
        public int prioridad { get; set; }
        public string estado_periodo { get; set; }
        public string comentario_revision_medida { get; set; }
        
        public DateTime fh_hasta { get; set; }
        public string subestado_SAP { get; set; }
        public string cod_subestado_SAP { get; set; }
        public DateTime fh_act { get; set; }
        public DateTime fh_prev_fin_crto { get; set; }
        public DateTime fh_baja_crto { get; set; }
        public string temporal { get; set; }
        public string ESTADO_GLOBAL { get; set; }
        public string ESTADO_GLOBAL_A_REPORTAR { get; set; }
        public string comentario { get; set; }
        public DateTime fecha_modificacion { get; set; }
    }
}
