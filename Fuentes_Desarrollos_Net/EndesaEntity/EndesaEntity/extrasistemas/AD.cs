﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.extrasistemas
{
    public class AD
    {
        public string empresa_emisora { get; set; }
        public string distribuidora { get; set; }
        public string cups22 { get; set; }
        public string cnae { get; set; }
        public string ind_activacion { get; set; }
        public DateTime fecha_activacion { get; set; }
        public string tarifa { get; set; }

        public int[] potencia { get; set; }
        public string tipo_identificador { get; set; }
        public string n_identificador { get; set; }
        public string tipo_persona { get; set; }
        public string razon_social { get; set; }
        public string telefono { get; set; }
        public string indicador_tipo_direccion { get; set; }
        public string direccion { get; set; }
        public string contacto { get; set; }
        public string tlf_contacto { get; set; }
        public string tipo_contrato { get; set; }
        public DateTime fecha_baja { get; set; }
        public string tipo_doc_aportado { get; set; }
        public string direccion_url { get; set; }
        public string observaciones { get; set; }

        public AD()
        {
            potencia = new int[6];
        }


    }
}
