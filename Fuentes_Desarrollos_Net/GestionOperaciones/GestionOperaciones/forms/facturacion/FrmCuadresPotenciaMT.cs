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
    public partial class FrmCuadresPotenciaMT : Form
    {
        EndesaBusiness.facturacion.InformeCuadrePotenciasMT cp;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmCuadresPotenciaMT()
        {
            usage.Start("Facturación", "FrmCuadresPotenciaMT" ,"N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {

            cp =
                new EndesaBusiness.facturacion.InformeCuadrePotenciasMT(txt_fd.Value, txt_fh.Value, txt_cups20.Text, txt_cnifdnic.Text);
            lbl_total_registros.Text = string.Format("Total Registros: {0:#,##0}", cp.listaFacturas.Count);
            dgv.AutoGenerateColumns = false;
            dgv.DataSource = cp.listaFacturas;
        }

        private void FrmCuadresPotenciaMT_Load(object sender, EventArgs e)
        {
            txt_fd.Text = new DateTime(2017, 01, 01).ToString();
            txt_fh.Text = new DateTime(2020, 12, 31).ToString();
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog save;
            save = new SaveFileDialog();
            save.Title = "Ubicación del archivo";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;
                cp.ExportarExcel(save.FileName);
                Cursor = Cursors.Default;
                MessageBox.Show("Informe terminado.",
                           "Exportación a Excel",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Information);
            }

        }

        private void FrmCuadresPotenciaMT_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmCuadresPotenciaMT" ,"N/A");
        }
    }
}
