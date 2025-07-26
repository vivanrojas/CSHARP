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
    public partial class FrmMes13ListaNegraCUPS_Edit : Form
    {
        public EndesaBusiness.factoring.Lista_Negra_Cups lista_negra { get; set; }

        public FrmMes13ListaNegraCUPS_Edit()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DateTime temp;

            if (txt_nif.Text == null || txt_nif.Text == "")
                errorProvider.SetError(txt_nif, "El campo debe estar informado.");
            else if (txt_cliente.Text == null || txt_cliente.Text == null)
                errorProvider.SetError(txt_cliente, "El campo debe estar informado.");
            else if (txt_cliente.Text == null || txt_cliente.Text == null)
                errorProvider.SetError(txt_cups20, "El campo debe estar informado.");
            else if (!DateTime.TryParse(txt_fecha_inicio.Text, out temp))
            {
                errorProvider.SetError(txt_fecha_inicio, "Fecha no válida.");
                
            }
            else if (!DateTime.TryParse(txt_fecha_fin.Text, out temp))
            {
                errorProvider.SetError(txt_fecha_fin, "Fecha no válida.");

            }
            else
            {
                lista_negra.nif = txt_nif.Text;
                lista_negra.cliente = txt_cliente.Text;
                lista_negra.cups20 = txt_cups20.Text;
                lista_negra.fecha_inicio = Convert.ToDateTime(txt_fecha_inicio.Text);
                lista_negra.fecha_fin = Convert.ToDateTime(txt_fecha_fin.Text);
                if(txt_comentario.Text != null)
                {
                    lista_negra.motivo = txt_comentario.Text;
                }
                lista_negra.Save();
                this.Close();

            }
        }
    }
}
