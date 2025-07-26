using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.medida
{
    public partial class FrmAdifCierresEnergia_Edit : Form
    {

        public EndesaBusiness.adif.CierresEnergia cierres { get; set; }
        public FrmAdifCierresEnergia_Edit()
        {
            InitializeComponent();
        }

        private void FrmAdifCierresEnergia_Edit_Load(object sender, EventArgs e)
        {

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (txt_cups20.Text == null || txt_cups20.Text == "")
                errorProvider.SetError(txt_cups20, "El campo debe estar informado.");
            else if (txt_fd.Value == null || txt_fd.Value == null)
                errorProvider.SetError(txt_fd, "El campo debe estar informado.");
            else if (txt_fh.Value == null || txt_fh.Value == null)
                errorProvider.SetError(txt_fd, "El campo debe estar informado.");
            else
            {
                cierres.cups20 = txt_cups20.Text;
                cierres.fecha_desde = txt_fd.Value;
                cierres.fecha_hasta= txt_fh.Value;
                cierres.Save();
                this.Close();

            }

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {

        }
    }
}
