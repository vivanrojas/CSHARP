using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class PNT_XML
    {
        public string codigoDelProceso { get; set; }
        public string codigoDePaso { get; set; }
        public int codigoDeSolicitud { get; set; }
        public string numExpediente { get; set; }
        public string cups { get; set; }
        public DateTime fechaSolicitud { get; set; }
        public DateTime fechaInspeccion { get; set; }
        public DateTime fechaAltaExpediente { get; set; }
        public string comentarios { get; set; }
        public string descripcion { get; set; }
        public string archivo { get; set; }
        public List<PNT_XML_Detail> lista_potencias { get; set; }
        public string factorCorreccion { get; set; }
        public DateTime fechaNormalizacion { get; set; }

        public PNT_XML()
        {
            lista_potencias = new List<PNT_XML_Detail>();
        }


    }
}
