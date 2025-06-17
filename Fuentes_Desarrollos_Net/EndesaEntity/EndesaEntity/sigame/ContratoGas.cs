using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.sigame
{
    public class ContratoGas
    {
        public int id_ps { get; set; }
        public int id_cto_ps { get; set; }
        public string nombre_cliente { get; set; }
        public string descripcion_ps { get; set; }
        public string nif { get; set; }
        public DateTime fecha_inicio { get; set; }
        public DateTime fecha_fin { get; set; }
        public int id_estado_contrato { get; set; }
        public string tarifa { get; set; }
        public int grupo_presion { get; set; }
        public string cupsree { get; set; }
        public string distribuidora { get; set; }
        public double qd { get; set; }
        public bool es_cisterna { get; set; }
        public string pais { get; set; }

        /* ADDED DATOS SUMINISTRO */
        public string contrato { get; set; }
        public string direccion { get; set; }
        public string numero_finca { get; set; }
        public string codigo_municipio { get; set; }
        public string id_provincia { get; set; }            
        public string codigo_postal { get; set; }        

        /* END ADDED DATOS SUMINISTRO */

        // [08/04/2025 GUS] Añadimos nuevo campo codigo_atr, relacionado con el código de la comercializadora en SIGAME y usado para rellenar el campo dispatchingcompany del XML generado en los adicionales
        public string codigo_atr {  get; set; }

    }
}
