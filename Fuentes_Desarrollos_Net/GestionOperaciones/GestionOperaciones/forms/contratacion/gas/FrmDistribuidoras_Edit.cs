using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.contratacion.gas
{
    
    public partial class FrmDistribuidoras_Edit : Form
    {
        public bool nuevo_registro { get; set; }
        EndesaBusiness.contratacion.gestionATRGas.Distribuidoras distri { get; set; }

        public bool changed { get; set; }
        public FrmDistribuidoras_Edit()
        {

            distri = new EndesaBusiness.contratacion.gestionATRGas.Distribuidoras();
            InitializeComponent();

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            bool todoOK = true;
            DateTime temp = new DateTime();

            if (txt_codigo.Text == null || txt_codigo.Text == "")
            {
                errorProvider.SetError(txt_codigo, "El código de distribuidora no puede estar en blanco.");
                todoOK = false;
            }
                
            if (txt_nombre.Text == null || txt_nombre.Text == "")
            {
                errorProvider.SetError(txt_nombre, "El nombre de distribuidora no puede estar en blanco.");
                todoOK = false;
            }

            if (!DateTime.TryParse(txt_fecha_desde.Text, out temp))
            {
                errorProvider.SetError(txt_fecha_desde, "Fecha no válida.");
                todoOK = false;
            }

            if (!DateTime.TryParse(txt_fecha_hasta.Text, out temp))
            {
                errorProvider.SetError(txt_fecha_hasta, "Fecha no válida.");
                todoOK = false;
            }            
                
            if (!txt_email.Text.Contains("@") && !txt_email.Text.Contains("."))
            {
                errorProvider.SetError(txt_email, "El campo eMail no tiene un formato válido.");
                todoOK = false;
            }              
            
            if(todoOK)
            {
                distri.codigo = txt_codigo.Text;
                distri.nombre = txt_nombre.Text;
                distri.fecha_desde = Convert.ToDateTime(txt_fecha_desde.Text);
                distri.fecha_hasta = Convert.ToDateTime(txt_fecha_hasta.Text);
                distri.email = txt_email.Text;

                if (nuevo_registro)
                    distri.Add();
                else
                    distri.Update();                    
                this.Close();
            }
                

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txt_codigo_TextChanged(object sender, EventArgs e)
        {
            changed = true;
        }

        private void txt_nombre_TextChanged(object sender, EventArgs e)
        {
            changed = true;
        }

        private void txt_fecha_desde_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void txt_fecha_desde_TextChanged(object sender, EventArgs e)
        {
            changed = true;
        }

        private void txt_fecha_hasta_TextChanged(object sender, EventArgs e)
        {
            changed = true;
        }

        private void txt_email_TextChanged(object sender, EventArgs e)
        {
            changed = true;
        }

        private void FrmDistribuidoras_Edit_Load(object sender, EventArgs e)
        {

        }
    }
}
