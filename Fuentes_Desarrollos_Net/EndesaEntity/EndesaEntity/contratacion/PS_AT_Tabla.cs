using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion
{
    public class PS_AT_Tabla
    {
        public string empresa { get; set; }
        public string cif { get; set; }
        public string nombre_cliente { get; set; }
        public string cups13 { get; set; }
        public string cups22 { get; set; }
        public string cups20 { get; set; }

        public string distribuidora { get; set; }
        public long contrato_atr { get; set; }
        public int cnumcatr { get; set; }
        public string version { get; set; }
        public int tdischor { get; set; }
        public string contrext { get; set; }
        public DateTime fecha_alta_contrato { get; set; }
        public DateTime fecha_baja_contrato { get; set; }
        public DateTime fecha_prevista_baja { get; set; }
        public int estado_contrato_id { get; set; }
        public string estado_contrato_descripcion { get; set; }
        public DateTime fecha_prevista_alta { get; set; }
        public string tarifa { get; set; }
        public DateTime fecha_puesta_en_servicio { get; set; }

        public string tindgcpy { get; set; }
        public string tticonps { get; set; }

        public string provincia { get; set; }
        public string segmentoMercado { get; set; }
        public int tension { get; set; }
        public double[] vpotcal { get; set; }
        public string tipoCli { get; set; }
        public int tipoGestionAtr { get; set; }

        public int tpuntmed { get; set; }
        public string descripcion_autoconsumo { get; set; }

        public DateTime fecha_anexion { get; set; }

        public DateTime min_fecha_anexion { get; set; }
        public DateTime max_fecha_anexion { get; set; }

        public bool migrado_sap { get; set; }

        public PS_AT_Tabla()
        {
            vpotcal = new double[6];
        }




    }
}
