using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.medida
{
    public partial class FrmLicitaciones : Form
    {
        public FrmLicitaciones()
        {
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void extraccionesMasivasEnFormatoP01011ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.medida.licitaciones.Licitaciones li = new EndesaBusiness.medida.licitaciones.Licitaciones();
            li.Extraccion();
        }
    }
}
