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
    public partial class FrmFacturasBTN_CAPB : Form
    {

        EndesaBusiness.facturacion.BTN_CAPB btn;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmFacturasBTN_CAPB()
        {
            DateTime fd;
            DateTime fh;
            DateTime mesAnterior;
            int anio;
            int mes;
            int dias_del_mes;

            usage.Start("Facturación", "FrmFacturasBTN_CAPB" ,"N/A");

            InitializeComponent();
            mesAnterior = DateTime.Now.AddMonths(-1);
            anio = mesAnterior.Year;
            mes = mesAnterior.Month;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            fh = new DateTime(anio, mes, dias_del_mes);
            fd = new DateTime(fh.Year, fh.Month, 1);

            this.txt_fecha_factura_desde.Value = fd;
            this.txt_fecha_factura_hasta.Value = fh;
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "lpc_btn_param";
            p.esquemaString = "FAC";
            p.cabecera = "Parámetros Cálculo Prefacturas BTN";
            p.Show();
        }

        private void btn_buscar_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            btn.Facturas(txt_fecha_factura_desde.Value, txt_fecha_factura_hasta.Value, txt_cups20.Text, chk_ultima_factura_periodo.Checked);
            lbl_total.Text = string.Format("Total Registros: {0:#,##0}", btn.lista.Count());
            dgv.AutoGenerateColumns = false;
            dgv.DataSource = btn.lista;
            Cursor.Current = Cursors.Default;
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            if (btn.lista.Count > 0)
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
                    btn.ExportExcel(save.FileName);
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

        private void FrmFacturasBTN_CAPB_Load(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            chk_ultima_factura_periodo.Checked = true;
            btn = new EndesaBusiness.facturacion.BTN_CAPB();
            Cursor.Current = Cursors.Default;
        }

        private void FrmFacturasBTN_CAPB_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmFacturasBTN_CAPB" ,"N/A");
        }
    }
}
