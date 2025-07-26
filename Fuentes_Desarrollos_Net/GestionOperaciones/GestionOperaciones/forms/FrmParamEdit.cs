using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms
{
    public partial class FrmParamEdit : Form
    {
        public FrmParamEdit()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (txtCode.Text == null || txtCode.Text == "")
            {
                MessageBox.Show("El código no puede estar en blanco",
                "Error en la validación de datos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                txtCode.Focus();
            }else if (txtDescription.Text == null || txtDescription.Text == "")
            {
                MessageBox.Show("La Descripción no puede estar blanco",
                "Error en la validación de datos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                txtDescription.Focus();
            }else
            {
                Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void FrmParamEdit_Load(object sender, EventArgs e)
        {

        }
    }
}
