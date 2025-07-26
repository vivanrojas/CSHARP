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
    public partial class FrmFacturasBTN : Form
    {
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmFacturasBTN()
        {

            DateTime fd;
            DateTime fh;
            DateTime mesAnterior;
            int anio;
            int mes;
            int dias_del_mes;

            usage.Start("Facturación", "FrmFacturasBTN" ,"N/A");

            InitializeComponent();

            mesAnterior = DateTime.Now.AddMonths(-1);
            anio = mesAnterior.Year;
            mes = mesAnterior.Month;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            fh = new DateTime(anio, mes, dias_del_mes);
            fd = new DateTime(fh.Year, fh.Month, 1);

            this.txt_fd.Value = fd;
            this.txt_fh.Value = fh;
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            SaveFileDialog save;           


            save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            DialogResult result = save.ShowDialog();            
            if (result == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;
                EndesaBusiness.facturacion.FacturasBTN btn = new EndesaBusiness.facturacion.FacturasBTN();
                btn.InformeFacturasBTN(Convert.ToDateTime(txt_fd.Text), Convert.ToDateTime(txt_fh.Text),save.FileName);
                
                if (!btn.hayError)
                {
                    MessageBox.Show("Informe terminado.",
                    "Exportación a Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Error al procesar el informe",
                   "Exportación a Excel",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
                }
                    
            }
        }

        private void FrmFacturasBTN_Load(object sender, EventArgs e)
        {
           
        }

        private void importarExcelAgrupadasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            double percent = 0;
            int progreso = 0;
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = true;
            EndesaBusiness.facturacion.FacturasAgrupadas fag = new EndesaBusiness.facturacion.FacturasAgrupadas();
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;
                forms.FrmProgressBar pb = new forms.FrmProgressBar();
                pb.Text = "Importando Excels";
                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = d.FileNames.Count();

                foreach (string fileName in d.FileNames)
                {
                    progreso++;
                    percent = (progreso / Convert.ToDouble(d.FileNames.Count())) * 100;
                    pb.progressBar.Increment(1);
                    pb.progressBar.Value = progreso;
                    pb.txtDescripcion.Text = "Importando archivo " + fileName;
                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    pb.Refresh();
                    fag.CargaExcel(fileName);
                }
                pb.Close();
                Cursor = Cursors.Default;
                MessageBox.Show("Importación finalizada",
                       "Importación de facturas agrupadas",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Information);

            }

        }

        private void FrmFacturasBTN_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmFacturasBTN" ,"N/A");
        }
    }
}
