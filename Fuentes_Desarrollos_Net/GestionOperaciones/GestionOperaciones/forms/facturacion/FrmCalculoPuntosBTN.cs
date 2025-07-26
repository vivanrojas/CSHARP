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
    public partial class FrmCalculoPuntosBTN : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmCalculoPuntosBTN()
        {
            usage.Start("Facturación", "FrmCalculoPuntosBTN" ,"N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void importarPlantillaDePuntosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.facturacion.puntos_calculo_btn.ImportacionPlantillaExcel imp_excel = new EndesaBusiness.facturacion.puntos_calculo_btn.ImportacionPlantillaExcel();

            DialogResult result = DialogResult.Yes;

            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Seleccione archivos Excel de Tunel";
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                result = MessageBox.Show("¿Desea procesar los Excels seleccionado?",
                 "Importación puntos prefacturas BTN",
                 MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    Cursor.Current = Cursors.WaitCursor;

                    foreach (string fileName in d.FileNames)
                    {
                        imp_excel.CargaExcel(fileName);
                    }

                    Cursor.Current = Cursors.Default;

                    MessageBox.Show("Proceso Finalizado."
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                       ,
                  "Importación ficheros Excel Tunel",
                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void rellenaPlantillaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.facturacion.puntos_calculo_btn.Calculo_Prefacturas_BTN calculo
                = new EndesaBusiness.facturacion.puntos_calculo_btn.Calculo_Prefacturas_BTN();
            calculo.Proceso();
        }

        private void FrmCalculoPuntosBTN_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmCalculoPuntosBTN" ,"N/A");
        }
    }
}
