using System;
using System.Windows.Forms;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FormPuntosSofisticados_Edit : Form
    {
        public bool pressCancel = false;
        public FormPuntosSofisticados_Edit()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (txtCCOUNIPS.Text == null || txtCCOUNIPS.Text == "")
            {
                MessageBox.Show("El campo CCOUNIPS no puede estar en blanco",
                "Error en la validación de datos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                txtCCOUNIPS.Focus();
            }
            else if (txtFD.Text == null)
            {
                MessageBox.Show("El campo Fecha Desde no puede estar en blanco",
                "Error en la validación de datos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                txtFD.Focus();
            }
            else
            {
                Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            pressCancel = true;
            Close();
        }
    }
}
