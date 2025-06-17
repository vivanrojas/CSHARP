using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.facturacion
{
    public class InformeRetrasoFacturacion
    {
        public String cnifdnic { get; set; }
        public String dapersoc { get; set; }
        public String tipoNegocio { get; set; }

        public List<MesDatosRetrasoFacturacion> listameses;
        public InformeRetrasoFacturacion()
        {
            listameses = new List<MesDatosRetrasoFacturacion>();
        }
    }
}
