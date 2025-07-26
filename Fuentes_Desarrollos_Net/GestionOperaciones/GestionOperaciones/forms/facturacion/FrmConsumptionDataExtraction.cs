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
    public partial class FrmConsumptionDataExtraction : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmConsumptionDataExtraction()
        {

            DateTime fd;
            DateTime fh;
            DateTime mesAnterior;
            int anio;
            int mes;
            int dias_del_mes;


            usage.Start("Facturación", "FrmConsumptionDataExtraction" ,"N/A");
            InitializeComponent();

            mesAnterior = DateTime.Now.AddMonths(-1);
            anio = mesAnterior.Year;
            mes = mesAnterior.Month;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            fh = new DateTime(anio, mes, dias_del_mes);
            fd = new DateTime(anio, 1, 1);

            txt_fecha_factura_desde.Value = fd;
            txt_fecha_factura_hasta.Value = fh;
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {

            EndesaBusiness.facturacion.ConsumptionDataExtration cde =
                new EndesaBusiness.facturacion.ConsumptionDataExtration(Convert.ToDateTime(txt_fecha_factura_desde.Text),
                Convert.ToDateTime(txt_fecha_factura_hasta.Text));


            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                cde.GeneraExcel(save.FileName);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }
        }

        private void FrmConsumptionDataExtraction_Load(object sender, EventArgs e)
        {
             
        }

        private void FrmConsumptionDataExtraction_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmConsumptionDataExtraction" ,"N/A");
        }
    }
}
