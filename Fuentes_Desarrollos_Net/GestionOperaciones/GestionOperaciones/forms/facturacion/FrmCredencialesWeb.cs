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
    public partial class FrmCredencialesWeb : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmCredencialesWeb()
        {

            usage.Start("Facturación", "FrmCredencialesWeb" ,"N/A");
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (txtDominio.Text == null || txtDominio.Text == "")
            {
                MessageBox.Show("El código no puede estar en blanco",
                "Error en la validación de datos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                txtDominio.Focus();
            }else if (txtUser.Text == null || txtUser.Text == "")
            {
                MessageBox.Show("El CodIntegr no puede ser blanco",
                "Error en la validación de datos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                txtUser.Focus();
            }else if (txtPass.Text == null || txtPass.Text == "")
            {
                MessageBox.Show("La Descripción no puede estar blanco",
                "Error en la validación de datos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                txtPass.Focus();
            }else
            {
                Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void txtCodIntegr_TextChanged(object sender, EventArgs e)
        {

        }

        private void FrmCredencialesWeb_Load(object sender, EventArgs e)
        {

        }

        private void FrmCredencialesWeb_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmCredencialesWeb" ,"N/A");
        }
    }
}
