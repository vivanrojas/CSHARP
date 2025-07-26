using System;
using System.Collections;
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
    public partial class FrmComprobacion_Contratacion : Form
    {

        EndesaBusiness.contratacion.PS_AT_HIST ps_at_hist;
        public FrmComprobacion_Contratacion()
        {
            InitializeComponent();
        }

        private void FrmComprobacion_Contratacion_Load(object sender, EventArgs e)
        {
            ps_at_hist = new EndesaBusiness.contratacion.PS_AT_HIST();
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void LoadData()
        {
            

            dgv_ps_at_hist.AutoGenerateColumns = false;
            dgv_ps_at_hist.DataSource = ps_at_hist.PS_AT_HIST_CUPS(txt_cups20.Text);

            //dgv_solicitudes_xxi.AutoGenerateColumns = false;
            //dgv_solicitudes_xxi.DataSource = lista;
        }
    }
}
