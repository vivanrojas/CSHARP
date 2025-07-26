using EndesaBusiness.contratacion;
using EndesaBusiness.contratacion.eexxi;
using EndesaBusiness.sharePoint;
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
    public partial class FrmFacturasBTE_CAPB : Form
    {

        EndesaBusiness.facturacion.BTE_CAPB bte;
        EndesaBusiness.utilidades.Global g = new EndesaBusiness.utilidades.Global();
        DateTime begin;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmFacturasBTE_CAPB()
        {
            DateTime fd;
            DateTime fh;
            DateTime mesAnterior;
            int anio;
            int mes;
            int dias_del_mes;

            usage.Start("Facturación", "FrmFacturasBTE_CAPB" ,"N/A");

            InitializeComponent();
            mesAnterior = DateTime.Now.AddMonths(-1);
            anio = mesAnterior.Year;
            mes = mesAnterior.Month;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            fh = new DateTime(anio, mes, dias_del_mes);
            fd = new DateTime(fh.Year, fh.Month, 1);

            //this.txt_fecha_factura_desde.Value = fd;
            //this.txt_fecha_factura_hasta.Value = fh;
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_buscar_Click(object sender, EventArgs e)
        {

            DateTime temp;
            bool todoOK = true;

            if (todoOK)
                if (!DateTime.TryParse(txt_fd_emision.Text, out temp) && !DateTime.TryParse(txt_fd_consumo.Text, out temp))
                {
                    errorProvider.SetError(txt_fd_emision, "Debe informar alguna fecha");
                    errorProvider.SetError(txt_fd_consumo, "Debe informar alguna fecha");
                    todoOK = false;
                }

            if (todoOK)
                if (!DateTime.TryParse(txt_fh_emision.Text, out temp) && !DateTime.TryParse(txt_fh_consumo.Text, out temp))
                {
                    errorProvider.SetError(txt_fh_emision, "Debe informar alguna fecha");
                    errorProvider.SetError(txt_fh_consumo, "Debe informar alguna fecha");
                    todoOK = false;
                }

            if (todoOK)
                if (DateTime.TryParse(txt_fd_emision.Text, out temp) && !DateTime.TryParse(txt_fh_emision.Text, out temp))
                {
                
                    errorProvider.SetError(txt_fh_emision, "Debe informar la fecha hasta");
                    todoOK = false;
                }

            if (todoOK)
                if (DateTime.TryParse(txt_fd_consumo.Text, out temp) && !DateTime.TryParse(txt_fh_consumo.Text, out temp))
                {

                    errorProvider.SetError(txt_fh_consumo, "Debe informar la fecha hasta");
                    todoOK = false;
                }

            if (todoOK)
                if (!DateTime.TryParse(txt_fd_emision.Text, out temp) && DateTime.TryParse(txt_fh_emision.Text, out temp))
                {

                    errorProvider.SetError(txt_fd_emision, "Debe informar la fecha desde");
                    todoOK = false;
                }

            if (todoOK)
                if (!DateTime.TryParse(txt_fd_consumo.Text, out temp) && DateTime.TryParse(txt_fh_consumo.Text, out temp))
                {

                    errorProvider.SetError(txt_fd_consumo, "Debe informar la fecha desde");
                    todoOK = false;
                }





            if (todoOK)
            {
                Cursor.Current = Cursors.WaitCursor;
                bte.Facturas(txt_fd_emision.Text, txt_fh_emision.Text,
                    txt_fd_consumo.Text, txt_fh_consumo.Text,
                    txt_cups20.Text);
                lbl_total.Text = string.Format("Total Registros: {0:#,##0}", bte.lista.Count());
                dgv.AutoGenerateColumns = false;
                dgv.DataSource = bte.lista;
                Cursor.Current = Cursors.Default;
            }

           
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {

            if(bte.lista.Count > 0)
            {
                SaveFileDialog save = new SaveFileDialog();
                save.Title = "Ubicación del informe";
                save.AddExtension = true;
                save.DefaultExt = "xlsx";
                save.Filter = "Ficheros xslx (*.xlsx)|*.*";
                DialogResult result = save.ShowDialog();
                if (result == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    bte.ExportExcel(save.FileName);
                    Cursor.Current = Cursors.Default;

                    DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (result2 == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(save.FileName);
                    }
                }
            }else
                MessageBox.Show("No hay datos para exportar a Excel.", "Abrir Excel generado",
                           MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void FrmFacturasBTE_CAPB_Load(object sender, EventArgs e)
        {
            bte = new EndesaBusiness.facturacion.BTE_CAPB();
            Cursor.Current = Cursors.WaitCursor;
            bte.Prefacturas();
            lbl_total_prefacturas.Text = string.Format("Total Registros: {0:#,##0}", bte.lista_prefacturas.Count());
            dgv_prefacturas.AutoGenerateColumns = false;
            dgv_prefacturas.DataSource = bte.lista_prefacturas;
            Cursor.Current = Cursors.Default;
        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void btn_Excel_Prefacturas_Click(object sender, EventArgs e)
        {
            if (bte.lista_prefacturas.Count > 0)
            {
                SaveFileDialog save = new SaveFileDialog();
                save.Title = "Ubicación del informe";
                save.AddExtension = true;
                save.DefaultExt = "xlsx";
                save.Filter = "Ficheros xslx (*.xlsx)|*.*";
                DialogResult result = save.ShowDialog();
                if (result == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    bte.ExportExcelPrefacturas(save.FileName);
                    Cursor.Current = Cursors.Default;

                    DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (result2 == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(save.FileName);
                    }
                }
            }
            else
                MessageBox.Show("No hay datos para exportar a Excel.", "Abrir Excel generado",
                           MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void txt_fd_consumo_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {
            errorProvider.Clear();
        }

        private void txt_fd_consumo_TextChanged(object sender, EventArgs e)
        {
            errorProvider.Clear();
        }

        private void txt_fh_consumo_TextChanged(object sender, EventArgs e)
        {
            errorProvider.Clear();
        }

        private void txt_fd_emision_TextChanged(object sender, EventArgs e)
        {
            errorProvider.Clear();
        }

        private void txt_fh_emision_TextChanged(object sender, EventArgs e)
        {
            errorProvider.Clear();
        }

        private void FrmFacturasBTE_CAPB_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmFacturasBTE_CAPB" ,"N/A");
        }
    }
}
