using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion
{
    public class SolATRMT
    {
        public int empresa_titular { get; set; }
        public string estadoSolAtr { get; set; }
        public long codSolAtr { get; set; }
        public string tipoSolAtr { get; set; }
        public int fRecepcion { get; set; }
        public int fAcepRech { get; set; }
        public int fRechazo { get; set; }
        public int fEnvioAtr { get; set; }
        public int fEnvioDoc { get; set; }
        public int fPrevAlta { get; set; }
        public int fActivacion { get; set; }
        public string ccounips { get; set; }
        public string cups_ext { get; set; }
        public string tipo_cli { get; set; }
        public string cif { get; set; }
        public string cliente { get; set; }
        public string seg_mer { get; set; }
        public int linea_n { get; set; }
        public string direc_ps { get; set; }
        public string cdistrib { get; set; }
        public long contr_atr { get; set; }
        public string coment_1_solicitud { get; set; }
        public string coment_2_solicitud { get; set; }
        public string coment_3_solicitud { get; set; }
        public string check_anulac { get; set; }
        public string mot_baja { get; set; }
        public double[] potencia { get; set; }
        public string tarifa { get; set; }
        public int tension { get; set; }
        public string ccontrps { get; set; }
        public int ver_contr_ps { get; set; }
        public string uso_contr { get; set; }
        public string gestor_sce { get; set; }
        public string coment_1_pettrab { get; set; }
        public string coment_2_pettrab { get; set; }
        public string coment_3_pettrab { get; set; }

        public string[] mot_rech { get; set; }

        public string[] coment_1_motrech { get; set; }
        public string[] coment_2_motrech { get; set; }
        public string[] coment_3_motrech { get; set; }

        public int testcont { get; set; }
        public string manual { get; set; }
        public string med_baj { get; set; }
        public long pt_cont { get; set; }
        public double prc_perd { get; set; }
        public string mod_fecha_solct { get; set; }
        public int fecha_solct { get; set; }
        public string mod_fecha_resp { get; set; }
        public int fecha_resp { get; set; }


        public SolATRMT()
        {
            potencia = new double[7];
            mot_rech = new string[10];
            coment_1_motrech = new string[10];
            coment_2_motrech = new string[10];
            coment_3_motrech = new string[10];
        }

    }
}
