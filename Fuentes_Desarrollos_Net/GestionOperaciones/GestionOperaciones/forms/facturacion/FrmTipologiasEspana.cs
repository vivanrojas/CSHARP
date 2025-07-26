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
    public partial class FrmTipologiasEspana : Form
    {
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmTipologiasEspana()
        {
            usage.Start("Facturación", "FrmTipologiasEspana" ,"N/A");
            InitializeComponent();
        }

        private void FrmTipologiasEspana_Load(object sender, EventArgs e)
        {

        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización Tipologías España";
            p.tabla = "irf_param";
            p.esquemaString = "CON";
            p.Show();
        }

        private void FrmTipologiasEspana_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmTipologiasEspana" ,"N/A");
        }
    }
}
