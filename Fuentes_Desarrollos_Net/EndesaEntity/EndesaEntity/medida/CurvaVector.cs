using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class CurvaVector
    {

        public double[] horariaActiva { get; set; }
        public double[] horariaReactiva { get; set; }
        public double[] cuartoHorariaActiva { get; set; }
        public double[] cuartoHorariaReactiva { get; set; }
        public DateTime[] cuartoHorariaFechaHora { get; set; }
        public double [] cuartoHorariaPotencia { get; set; }
        public int numPeriodosHorarios { get; set; }
        public int numPeriodosCuartoHorarios { get; set; }

        public CurvaVector(int _numPeriodosHorarios)
        {
            numPeriodosHorarios = _numPeriodosHorarios;
            numPeriodosCuartoHorarios = numPeriodosHorarios * 4;
            cuartoHorariaActiva = new double[numPeriodosCuartoHorarios + 1];
            cuartoHorariaReactiva = new double[numPeriodosCuartoHorarios + 1];
            cuartoHorariaFechaHora = new DateTime[numPeriodosCuartoHorarios + 1];
        }

    }
}
