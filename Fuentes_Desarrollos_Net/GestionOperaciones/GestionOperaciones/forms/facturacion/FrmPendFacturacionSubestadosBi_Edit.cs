using EndesaBusiness.factoring;
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
    public partial class FrmPendFacturacionSubestadosBi_Edit : Form
    {
        public EndesaBusiness.facturacion.redshift.Pendiente_Subestados pendiente { get; set; }
        public FrmPendFacturacionSubestadosBi_Edit()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (txt_cod_subestado.Text == null || txt_cod_subestado.Text == "")
                errorProvider.SetError(txt_cod_subestado, "El campo debe estar informado.");
            else if (txt_subestado_descripcion.Text == null || txt_subestado_descripcion.Text == null)
                errorProvider.SetError(txt_subestado_descripcion, "El campo debe estar informado.");

            else
            {
                pendiente.cod_subestado = txt_cod_subestado.Text;
                pendiente.descripcion_subestado = txt_subestado_descripcion.Text;
                pendiente.area_responsable = cmb_area_responsable.Text;
                pendiente.Save();
                this.Close();

            }
        }
    }
}
