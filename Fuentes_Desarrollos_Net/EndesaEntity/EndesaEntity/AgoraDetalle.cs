using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class AgoraDetalle
    {
        public int idps { get; set; } // para gas
        public DateTime fd { get; set; }
        public DateTime fh { get; set; }
        public int fecha_medida { get; set; }
        public string grupo { get; set; }
        public string grupo_informe { get; set; }
        public string nif { get; set; }
        public string nombre_cliente { get; set; }
        public string cups13 { get; set; }
        public string cups20 { get; set; }
        public string estado_contrato { get; set; }
        public string ultimo_mes_facturado { get; set; } // yyyymm
        public string estado_ltp { get; set; }
        public string tipo_informe { get; set; } // Nos indica si es Agora SCE, Sofisticado, etc ...
        public double tam { get; set; } // promedio facturacion
        public string tipo_tam { get; set; }
        public string comentario { get; set; }
        public string informe { get; set; }
        public string alarma { get; set; }
        public int aaaammPdte { get; set; }
        public string cod_contrato { get; set; }
        public string distribuidora { get; set; }
        public string empresa_titular { get; set; }
        public string estado_pendiente { get; set; }
        public string subestado_pendiente { get; set; }
        public string provincia { get; set; }
        public string precios { get; set; }
        public string tipo { get; set; }
        public bool facturado { get; set; }
        public bool existeMedidasyFacturas { get; set; }
        public bool existeMedidasyFacturasOK { get; set; }
        public bool existe_T_SGM_M_LECTURASyCONSUMOS { get; set; }
        public string cd_pais { get; set; }
        public int id_estado { get; set; }
        public int id_pmedida { get; set; }
        public bool esTop { get; set; }
        public bool tam_en_blanco { get; set; }
        public string nombreGestor { get; set; }
        public string apellido1 { get; set; }
        public string apellido2 { get; set; }
        public string desc_Responsable_Territorial { get; set; }
        public string subDireccion { get; set; }

        public AgoraDetalle()
        {
            tam_en_blanco = false;
        }

    }
}
