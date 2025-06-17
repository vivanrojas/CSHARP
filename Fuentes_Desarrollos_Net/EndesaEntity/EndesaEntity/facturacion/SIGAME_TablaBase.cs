using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class SIGAME_TablaBase
    {

        public int id_ps { get; set; }
        public int secfactu { get; set; }
        public string numrefFactura { get; set; }
        public string origen { get; set; }
        public string marca { get; set; }
        public int periodo { get; set; }
        public DateTime fh_ini_facturacion { get; set; }
        public DateTime fh_fin_facturacion { get; set; }
        public string periodofacturacion { get; set; }
        public string razon_social { get; set; }
        public string cifnif { get; set; }
        public string cups22 { get; set; }
        public DateTime fh_emision { get; set; }
        public string c_factura { get; set; }
        public string codificacion_factura { get; set; }
        public string tipologia_factura { get; set; }
        public string direccion_punto_suministro { get; set; }        
        public int consumo_total { get; set; }

        public double importe_ih { get; set; }
        public double base_imponible_iva { get; set; }
        public double nm_importe_iva { get; set; }
        public double ImporteTotalFactura { get; set; }
        public double ImporteTotalFactura_sce { get; set; }

        public string IHTI_descripcion { get; set; }
        public double IHTI_tipo_impositivo { get; set; }
        public int IHTI_consumo { get; set; }
        public double IHTI_importe { get; set; }

        public string IHTD_descripcion { get; set; }
        public double IHTD_tipo_impositivo { get; set; }
        public int IHTD_consumo { get; set; }
        public double IHTD_importe { get; set; }

        public string IHTC_descripcion { get; set; }
        public double IHTC_tipo_impositivo { get; set; }
        public int IHTC_consumo { get; set; }
        public double IHTC_importe { get; set; }

        public string uso_gas { get; set; }
        
        
        
        public List<EndesaEntity.facturacion.SIGAME_TablaBaseDetalle> l { get; set; }

        public SIGAME_TablaBase()
        {
            l = new List<SIGAME_TablaBaseDetalle>();
        }
    }
}
