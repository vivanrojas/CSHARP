using EndesaBusiness.contratacion;
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
    public partial class FrmPeajesFormatoF : Form
    {
        EndesaBusiness.medida.ExcelCUPS ex;
        List<EndesaEntity.medida.PuntoSuministro> lc = new List<EndesaEntity.medida.PuntoSuministro>();
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmPeajesFormatoF()
        {
            usage.Start("Medida", "FrmPeajesFormatoF", "N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
               

        private void GeneraExcel(List<EndesaEntity.medida.PuntoSuministro> lc, bool fecha_consumo)
        {
            EndesaBusiness.medida.Redshift.Peajes peajes = new EndesaBusiness.medida.Redshift.Peajes();

            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                peajes.ExportExcel(save.FileName, lc, fecha_consumo);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }

        }

       
        private void FrmPeajesFormatoF_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Medida", "FrmPeajesFormatoF", "N/A");
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = false;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    ex = new EndesaBusiness.medida.ExcelCUPS(fileName);
                    if (!ex.hayError)
                    {
                        lc = ex.lista_cups;
                        Carga_DGV(lc);
                    }
                    else
                    {
                        MessageBox.Show(ex.descripcion_error,
                           "Importar Excel",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Information);
                    }

                }
            }
        }

        private void Carga_DGV(List<EndesaEntity.medida.PuntoSuministro> lista)
        {
            lblTotalRegistros.Text = string.Format("Total Registros: {0:#,##0}", lista.Count);
            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;
        }

        private void btn_generar_excels_Click(object sender, EventArgs e)
        {
            GeneraExcel(lc, cmbTipoFecha.SelectedIndex == 1);
        }

        private void btnExcel_Click_1(object sender, EventArgs e)
        {
            if(txt_fh.Value < txt_fd.Value)
                errorProvider.SetError(txt_fh, "La fecha hasta no puede ser menor que la fecha desde.");
            if (txt_cups20.Text == null || txt_cups20.Text == "")
                errorProvider.SetError(txt_cups20, "El campo CUPS20 debe estar informado.");
            else if (txt_cups20.Text.Length != 20)
                errorProvider.SetError(txt_cups20, "El campo CUPS20 no tiene 20 caracteres.");
            else
            {
                lc.Clear();
                EndesaEntity.medida.PuntoSuministro c = new EndesaEntity.medida.PuntoSuministro();
                c.fd = txt_fd.Value;
                c.fh = txt_fh.Value;
                c.cups20 = txt_cups20.Text;
                lc.Add(c);
                GeneraExcel(lc, cmbTipoFecha.SelectedIndex == 1);
            }


        }

        private void FrmPeajesFormatoF_Load(object sender, EventArgs e)
        {
            DateTime mesAnterior = DateTime.Now.AddMonths(-1);
            int anio = mesAnterior.Year;
            int mes = mesAnterior.Month;
            int dias_del_mes = DateTime.DaysInMonth(anio, mes);

            DateTime fh = new DateTime(anio, mes, dias_del_mes);
            DateTime fd = new DateTime(fh.Year, fh.Month, 1);

            this.cmbTipoFecha.SelectedIndex = 0;
            this.cmbTipoFecha_masiva.SelectedIndex = 0;

            txt_fd.Value = fd;
            txt_fh.Value = fh;
        }
    }
}
