using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EndesaEntity.extrasistemas
{
    public class Global
    {
        public string hoja { get; set; }
        public string empresa_emisora { get; set; }
        public string distribuidora { get; set; }
        public string cups22 { get; set; }
        public string tipo_modificacion { get; set; }
        public string tipo_solicitud_administrativa { get; set; }
        public string cnae { get; set; }
        public string ind_activacion { get; set; }
        public DateTime fecha_activacion { get; set; }
        public DateTime fecha_baja { get; set; }
        public DateTime fecha_finalizacion { get; set; }

        public DateTime fechaAceptacion { get; set; }  //  irh 20/05/25
        public string contratacion_incondicional_ps { get; set; }        
        public string contratacion_incondicional_bs { get; set; }
        public string tipo_contrato { get; set; }
        public string tipo_contrato_atr { get; set; }
        public string tarifa_actual { get; set; }
        public int[] potencias_actual { get; set; }
        public string tarifa_nueva { get; set; }

        public string tarifa { get; set; }
        public int[] potencias { get; set; }
        public int[] potencias_nuevas { get; set; }

        public string tension_solicitada { get; set; }
        public string tipocups { get; set; }             //   irh 10/03/25 - tipo cups x(2) 
        public string cau { get; set; }                 //   irh 10/03/25 - Codigo de autoconsumo x(26)
        public string indesencial { get; set; }        //  irh   m101
        public string tipo_autoconsumo { get; set; }
        public string tipo_subseccion { get; set; }     //  irh  10/03/25   x(2)
        public string colectivo { get; set; }            //  irh  10/03/25   X(1)
        public int potinstaladagen { get; set; }     //  irh  10/03/25    9(14) 
        public string tipoinstalacion { get; set; }     //  irh  10/03/25   x(2)
        public string ssaa { get; set; }               //  irh  10/03/25   x(1)
        public string unicocontrato { get; set; }     //  irh  10/03/25   x(1)
        public string solicitud_modificacion_tension { get; set; }
        public string modo_control_potencia { get; set; }
        public int tension { get; set; }
        public string contacto { get; set; }
        public string tlf_contacto { get; set; }
        public string tipo_identificador { get; set; }
        public string n_identificador { get; set; } 
        public string tipo_persona { get; set; }
        public string razon_social { get; set; }
        public string nombre_de_pila { get; set; }
        public string primer_apellido { get; set; }
        public string telefono { get; set; }
        public string indicador_tipo_direccion { get; set; }
        public string direccion { get; set; }
        public List<Documentacion> lista_documentacion { get; set; }
        public List<string> direccion_url { get; set; }        
        public string g_empresarial { get; set; }
        public string entorno { get; set; }
        public string tipo_cliente { get; set; }
        public string tlf_contracto { get; set; }
        public string motivo_baja { get; set; }
        public string pais { get; set; }
        public string provincia { get; set; }
        public string codigo_ine_municipio { get; set; }
        public string codigo_postal { get; set; }
        public string persona_contacto { get; set; }      
        public string tipo_via { get; set; }
        public string numero { get; set; }
        public string piso { get; set; }
        public string puerta { get; set; }

        public string municipio { get; set; }
        public string nombre_via { get; set; }
        public string fichero { get; set; }
        public string observaciones { get; set; }
        
        public Global()
        {
            potencias = new int[6];
            potencias_actual = new int[6];
            potencias_nuevas = new int[6];
            lista_documentacion = new List<Documentacion>();
        }
    }

    public class Documentacion
    {
        public string tipo_doc_aportado { get; set; }
        public string direccion_url { get; set; }
    }
}
