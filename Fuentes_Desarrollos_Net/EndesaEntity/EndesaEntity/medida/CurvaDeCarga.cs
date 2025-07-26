using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class CurvaDeCarga
    {
        public  DateTime fecha { get; set; }
        public int numPeriodosHorarios { get; set; }
        public int numPeriodosCuartoHorarios { get; set; }
        public double[] horaria_activa { get; set; }
        public bool[] existe_horaria_activa { get; set; }
        public double[] horaria_reactiva { get; set; }
        public bool[] existe_horaria_reactiva { get; set; }
        public double[] cuartohoraria_activa { get; set; }
        public double[] cuartohoraria_reactiva { get; set; }
        public double[] potencias_cuartohorarias { get; set; }
        public double total_activa { get; set; }
        public double total_reactiva { get; set; }
        public string fuente { get; set; }
        
        public CurvaDeCarga()
        {
            horaria_activa = new double[25];
            existe_horaria_activa = new bool[25];
            horaria_reactiva = new double[25];
            existe_horaria_reactiva = new bool[25];
            cuartohoraria_activa = new double[101];
            potencias_cuartohorarias = new double[101];





        }
    }
}

