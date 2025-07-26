using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.xxi
{
    public class XML_Aparatos
    {
        public string tipoAparato { get; set; }
        public  string marcaAparato { get; set; }
        public string modeloMarca { get; set; }
        public string tipoMovimiento { get; set; }
        public string tipoEquipoMedida { get; set; }
        public string tipoPropiedadAparato { get; set; } 
        public string tipoDHEdM { get; set; }
        public string modoMedidaPotencia { get; set; }
        public string codPrecinto { get; set; }
        public string periodoFabricacion { get; set; }
        public string numeroSerie { get; set; }
        public string funcionAparato { get; set; }
        public string numIntegradores { get; set; }
        public string constanteEnergia { get; set; }
        public string constanteMaximetro { get; set; }
        public string ruedasEnteras { get; set; }
        public string ruedasDecimales { get; set; }
        public List<XML_Medidas> listaMedidas { get; set; }
        
        public XML_Aparatos()
        {
            listaMedidas = new List<XML_Medidas>();
        }

    }
}
