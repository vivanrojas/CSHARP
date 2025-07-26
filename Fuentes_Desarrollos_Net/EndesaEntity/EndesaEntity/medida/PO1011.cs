using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class PO1011
    {
        public string cups22 { get; set; }
        public int tipo_medida { get; set; }
        public DateTime fecha_hora { get; set; }
        public int estacion { get; set; }
        public double ae { get; set; }
        public bool ae_nulo { get; set; }
        public int cal_ae { get; set; }
        public double _as { get; set; }
        public int cal_as { get; set; }
        public double[] r { get; set; }
        public bool r_nulo { get; set; }
        public int[] cal_r { get; set; }
        public double mag_reserv_1 { get; set; }
        public int cal_mag_reserv_1 { get; set; }
        public double mag_reserv_2 { get; set; }
        public int cal_mag_reserv_2 { get; set; }
        public int metodo_obtencion { get; set; }
        public int indicador_firmeza { get; set; }
        public string archivo { get; set; }
        public int fecha_archivo { get; set; }


        public PO1011()
        {
            ae_nulo = false;
            r_nulo = false;
            r = new double[4];
            cal_r = new int[4];
        }
    }
}
