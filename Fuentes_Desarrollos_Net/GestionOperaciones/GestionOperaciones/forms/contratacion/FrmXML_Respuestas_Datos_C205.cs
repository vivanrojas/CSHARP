using EndesaEntity.cnmc.V21_2019_12_17;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.contratacion
{
       

    public partial class FrmXML_Respuestas_Datos_C205 : Form
    {
        public EndesaBusiness.cnmc.CNMC cnmc { get; set; }
     //   public EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 xml { get; set; }
    //   
        public EndesaBusiness.xml.ContratacionXML cont_xml { get; set; }
    //    public Dictionary<string, string> dic_tipo_activacion_prevista { get; set; }
        public List<string> lista_motivo_rechazo { get; set; }
        
        public DateTime fecha { get; set; }

        public FrmXML_Respuestas_Datos_C205()
        {
            InitializeComponent();
        }

         private void btnOK_Click(object sender, EventArgs e)
         {
                             

            this.DialogResult = DialogResult.OK;
            this.Close();

           
        }

               

        private void FrmXML_Respuestas_Fecha_Load(object sender, EventArgs e)
        {
           
        }

        private void chk_rechazo_CheckedChanged(object sender, EventArgs e)
        {
          
        }

        private void chk_aceptacion_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void txt_codcontrato_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }
    }
}
