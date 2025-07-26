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
    public partial class FrmIHGasEspana : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmIHGasEspana()
        {

            usage.Start("Facturación", "FrmIHGasEspana" ,"N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            EndesaBusiness.facturacion.FacturasGas fg = new EndesaBusiness.facturacion.FacturasGas();


            MessageBox.Show("Debido a la complejidad del informe y a la calidad"
                             + System.Environment.NewLine
                             + "de los datos, será muy importe revisar y comprobar"
                             + System.Environment.NewLine
                             + "los resultados."
                             + System.Environment.NewLine
                             + System.Environment.NewLine
                             + "Debido a la búsqueda de facturas tanto en SIGAME"
                             + System.Environment.NewLine
                             + "y SCE, el informe, para un año de búsqueda,"
                             + System.Environment.NewLine
                             + "tardará entre 5 y 10 minutos."
                             + System.Environment.NewLine
                             + "Por favor, sea paciente.",                             
                            "Aviso muy importarte.",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);


            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del archivo Excel";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                fg.InformeFiscalGas(txtFD.Value, txtFH.Value, save.FileName);
                MessageBox.Show("Informe generado","Finalización del informe",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
            }

                
        }

        private void FrmIHGasEspana_Load(object sender, EventArgs e)
        {
            txtFD.Value = new DateTime(DateTime.Now.Year - 1, 1, 1);
            txtFH.Value = new DateTime(DateTime.Now.Year - 1, 12, 31);
        }

        private void FrmIHGasEspana_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmIHGasEspana" ,"N/A");
        }
    }
}
