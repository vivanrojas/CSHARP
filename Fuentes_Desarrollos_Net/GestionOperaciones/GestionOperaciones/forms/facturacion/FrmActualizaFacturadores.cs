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
    public partial class FrmActualizaFacturadores : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmActualizaFacturadores()
        {

            usage.Start("Facturación", "FrmActualizaFacturadores" ,"N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_programas_consumo_Click(object sender, EventArgs e)
        {

            bool hay_error;
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = false;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    EndesaBusiness.facturacion.ActualizaFacturadores acfact =
                        new EndesaBusiness.facturacion.ActualizaFacturadores();

                    acfact.LoadDataFromExcelProgramas(fileName);
                    

                }
            }
        }

        private void FrmActualizaFacturadores_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.Start("Facturación", "FrmActualizaFacturadores" ,"N/A");
        }
    }
}
