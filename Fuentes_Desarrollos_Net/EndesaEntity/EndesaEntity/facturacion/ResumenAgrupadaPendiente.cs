using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class ResumenAgrupadaPendiente
    {
        public string cif { get; set; }
        public string razon_social { get; set; }
        public int num_dia_agrupacion { get; set; }
        public int pendiente_medida { get; set; }
        public int oc_calculable { get; set; }
        public int dc_generado_sin_di { get; set; }
        public int di_apartado { get; set; }
        public double tam_total { get; set; }
        public bool es_agora { get; set; } //Con que haya un PS que sea Agora para un CIF lo consideramos true
        
        public ResumenAgrupadaPendiente()
        {
            this.pendiente_medida = 0;
            this.oc_calculable = 0;
            this.dc_generado_sin_di = 0;
            this.di_apartado = 0;
            this.tam_total = 0;
            this.es_agora = false;
        }
    }
}
