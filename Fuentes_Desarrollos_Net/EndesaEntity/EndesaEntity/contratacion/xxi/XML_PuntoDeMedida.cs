using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.xxi
{
    public class XML_PuntoDeMedida
    {
        public string codPM { get; set; }
        public string tipoMovimiento { get; set; }
        public string tipoPM { get; set; }
        public string funcion { get; set; }
        public string telefonoTelemedida { get; set; }
        public string modoLectura { get; set; }
        public string tensionPM { get; set; }
        public DateTime fechaVigor { get; set; }
        public DateTime fechaAlta { get; set; }
        public DateTime fechaBaja { get; set; }
        public List<XML_Aparatos> lista_aparatos { get; set; }
        

        public XML_PuntoDeMedida()
        {
            lista_aparatos = new List<XML_Aparatos>();
        }
    }
}
