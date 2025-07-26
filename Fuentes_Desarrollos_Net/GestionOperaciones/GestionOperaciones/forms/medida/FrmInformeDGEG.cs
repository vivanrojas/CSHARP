using EndesaBusiness.punto_suministro;
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
    public partial class FrmInformeDGEG : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmInformeDGEG()
        {

            usage.Start("Medida", "FrmInformeDGEG", "N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            EndesaBusiness.medida.Informe_DGEG dgeg = new EndesaBusiness.medida.Informe_DGEG();
            
            SaveFileDialog save;
            save = new SaveFileDialog();
            save.Title = "Ubicación del archivo Excel";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;
                dgeg.Informe(save.FileName, txtFFACTURADES.Value, txtFFACTURAHAS.Value);
                Cursor = Cursors.Default;
                MessageBox.Show("Informe terminado.",
                    "Exportación a Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {
            txtFFACTURADES.Value = new DateTime(DateTime.Now.AddYears(-1).Year, 1, 1);
            txtFFACTURAHAS.Value = new DateTime(DateTime.Now.AddYears(-1).Year, 12, 31);            
        }

        private void FrmInformeDGEG_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Medida", "FrmInformeDGEG", "N/A");
        }
    }
}
