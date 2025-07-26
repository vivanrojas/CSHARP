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
    public partial class FrmMes13ListaNegra_Edit : Form
    {

        public EndesaBusiness.factoring.Lista_Negra lista_negra { get; set; }
        public FrmMes13ListaNegra_Edit()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {

            if (txt_nif.Text == null || txt_nif.Text == "")
                errorProvider.SetError(txt_nif, "El campo debe estar informado.");
            else if (txt_cliente.Text == null || txt_cliente.Text == null)
                errorProvider.SetError(txt_cliente, "El campo debe estar informado.");
           
            else
            {
                lista_negra.nif = txt_nif.Text;
                lista_negra.cliente = txt_cliente.Text;
                lista_negra.Save();
                this.Close();                

            }


            
        }
    }
}
