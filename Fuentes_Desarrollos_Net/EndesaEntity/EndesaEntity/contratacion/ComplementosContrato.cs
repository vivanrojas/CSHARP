using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion
{
    public class ComplementosContrato
    {
        public string cups13 { get; set; }
        public string codigo_complemento { get; set; }
        public string producto { get; set; }
        public double[] vparam { get; set; }
        public string ccompobj { get; set; }
        public string catalogo { get; set; }

        public ComplementosContrato()
        {
            vparam = new double[5];
        }
    }
}
