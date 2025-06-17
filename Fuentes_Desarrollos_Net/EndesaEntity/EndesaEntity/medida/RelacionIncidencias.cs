using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class RelacionIncidencias: Audit
    {
   
        public string cups { get; set; }
        public string Incidencia { get; set; }

        public string area_responsable { get; set; }

        public string periodo { get; set; }

        public string Estado_Fac_SE { get; set; }

        public string Titulo_FAC { get; set; }

    }
}
