using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmMes13Comprobacion : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmMes13Comprobacion()
        {
            usage.Start("Facturación", "FrmMes13Comprobacion", "N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmMes13Comprobacion_Load(object sender, EventArgs e)
        {

        }

        private void FrmMes13Comprobacion_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmMes13Comprobacion", "N/A");
        }
    }
}
