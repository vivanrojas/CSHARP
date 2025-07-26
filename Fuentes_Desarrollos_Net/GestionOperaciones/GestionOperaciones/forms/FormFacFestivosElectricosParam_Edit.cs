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
    public partial class FormFacFestivosElectricosParam_Edit : Form
    {
        public FormFacFestivosElectricosParam_Edit()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (txtFecha.Value == null)
            {
                MessageBox.Show("La fecha no puede estar en blanco",
                "Error en la validación de datos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                txtFecha.Focus();
            }else if (txtDescripcion.Text == null || txtDescripcion.Text == "")
            {
                MessageBox.Show("La Descripción no puede estar blanco",
                "Error en la validación de datos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                txtDescripcion.Focus();
            }else
            {
                Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
