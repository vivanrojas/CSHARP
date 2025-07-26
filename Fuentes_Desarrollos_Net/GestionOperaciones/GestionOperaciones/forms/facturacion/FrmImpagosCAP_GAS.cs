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
    public partial class FrmImpagosCAP_GAS : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        public FrmImpagosCAP_GAS()
        {

            usage.Start("Facturación", "FrmImpagosCAP_GAS", "N/A");

            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void procesarExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.facturacion.cap_gas.Calculo_CAP_GAS ex =
                new EndesaBusiness.facturacion.cap_gas.Calculo_CAP_GAS();

            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = false;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    ex.ExcelFacturas(fileName);
                    if (!ex.hayError)
                    {
                       
                    }
                    else
                    {
                        MessageBox.Show(ex.descripcion_error,
                           "Importar Excel",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error);
                    }

                }

                MessageBox.Show("Proceso Terminado",
                           "Importar Excel",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Information);
            }
        }

        private void procesarExcelDeCUPSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.facturacion.cap_gas.Calculo_CAP_GAS ex =
               new EndesaBusiness.facturacion.cap_gas.Calculo_CAP_GAS();

            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = false;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    ex.ExcelCUPS(fileName);
                    if (!ex.hayError)
                    {

                    }
                    else
                    {
                        MessageBox.Show(ex.descripcion_error,
                           "Importar Excel",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error);
                    }

                }

                MessageBox.Show("Proceso Terminado",
                           "Importar Excel",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Information);
            }
        }

        private void FrmImpagosCAP_GAS_Load(object sender, EventArgs e)
        {

        }

        private void FrmImpagosCAP_GAS_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmImpagosCAP_GAS", "N/A");
        }

        private void procesarExcelAgrupadasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.facturacion.cap_gas.Calculo_CAP_GAS ex =
               new EndesaBusiness.facturacion.cap_gas.Calculo_CAP_GAS();

            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = false;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    ex.ExcelAgrupadas(fileName);
                    if (!ex.hayError)
                    {

                    }
                    else
                    {
                        MessageBox.Show(ex.descripcion_error,
                           "Importar Excel",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error);
                    }

                }

                MessageBox.Show("Proceso Terminado",
                           "Importar Excel",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Information);
            }
        }
    }
}
