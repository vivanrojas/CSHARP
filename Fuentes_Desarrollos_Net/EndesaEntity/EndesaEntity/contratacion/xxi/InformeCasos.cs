using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.xxi
{
    public class InformeCasos
    {

        public string nif { get; set; }
        public string razon_social { get; set; }
        public string cups13 { get; set; }
        public string cups20 { get; set; }
        public string cups22 { get; set; }
        public string codigo_solicitud { get; set; }
        public DateTime fecha_activacion { get; set; } // ALTA
        public DateTime fecha_baja { get; set; } // BAJA Fecha de la activación de la baja
        public int caso { get; set; }
        public string tipo_xml { get; set; }
        public string descripcion { get; set; }
        public bool existe_baja { get; set; }
        public bool existe_ps { get; set; }
        public bool existe_ps_hist { get; set; }
        public string empresa { get; set; }
        public string estado_contrato_ps { get; set; }
        public bool crear_contrato { get; set; }
        public bool crear_incidencia { get; set; }
        public bool realizar_seguimiento { get; set; }
        public string acciones { get; set; }

        
       


    }
}
