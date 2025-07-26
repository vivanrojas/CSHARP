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
    
    public partial class FrmContratosDetalle_Edit : Form
    {
        public bool nuevo_registro { get; set; }
        public bool changed { get; set; }
        public string nif { get; set; }

        public string tarifa { get; set; }

        public EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle con { get; set; }
        public FrmContratosDetalle_Edit()
        {               

            InitializeComponent();

            // LoadData();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DateTime temp;
            bool todoOK = true;

            if (cmb_tipo.Text == null || cmb_tipo.Text == "")
            {
                errorProvider.SetError(cmb_tipo, "El campo debe estar informado.");
                todoOK = false;
            }           
                                               
            if (!DateTime.TryParse(txt_fecha_inicio.Text, out temp))
            {
                errorProvider.SetError(txt_fecha_inicio, "Fecha no válida.");
                todoOK = false;
            }
                
            // No se obliga que la fecha se quede en blanco
            //if (cmb_tipo.Text != "INDEFINIDO")
            //{
            //    if (!DateTime.TryParse(txt_fecha_fin.Text, out temp))
            //    {
            //        errorProvider.SetError(txt_fecha_fin, "Fecha no válida.");
            //        todoOK = false;
            //    }                    
            //}            
                        
            if (txt_qd.Text == null || txt_qd.Text == "")
            {
                errorProvider.SetError(txt_qd, "El campo debe estar informado.");
                todoOK = false;
            }


            if (con.ExisteRegistro(nif, txt_cups20.Text, Convert.ToDateTime(txt_fecha_inicio.Text), 
                cmb_tipo.SelectedItem.ToString(), Convert.ToDouble(txt_qd.Text.Replace(",",".")), txt_comentario.Text))
            {                

                MessageBox.Show("Ya existe un " + cmb_tipo.SelectedItem.ToString() + " para la fecha "
                       + Convert.ToDateTime(txt_fecha_inicio.Text).ToString("dd/MM/yyyy") + " ~ "
                       + (txt_fecha_fin.Text.Replace("/", "").Trim() != "" ? Convert.ToDateTime(txt_fecha_fin.Text).ToString("dd/MM/yyyy") : "")
                       + System.Environment.NewLine
                       + "Si desea incluir el mismo producto para las mismas fechas"
                       + System.Environment.NewLine
                       + "deberá insertar un comentario justificándolo.",
                       "Detectado producto repetido",
                      MessageBoxButtons.OK, MessageBoxIcon.Warning);

                errorProvider.SetError(txt_comentario, "El campo debe estar informado.");
                todoOK = false;
            }


            if (todoOK)
            {                
                con.cups20 = txt_cups20.Text;                
                con.fecha_inicio = Convert.ToDateTime(txt_fecha_inicio.Text);
                con.fecha_fin = txt_fecha_fin.Text.Replace("/","").Trim() != "" ? Convert.ToDateTime(txt_fecha_fin.Text) : DateTime.MinValue;

                //if (cmb_tipo.Text != "INDEFINIDO")
                //    con.fecha_fin = Convert.ToDateTime(txt_fecha_fin.Text);
                //else
                //    con.fecha_fin = DateTime.MinValue;

                con.nif = this.nif;
                con.qd = Convert.ToDouble(txt_qd.Text);
                con.tarifa = cmb_tarifa.SelectedItem == null ? "" : cmb_tarifa.SelectedItem.ToString();
                con.tipo = cmb_tipo.SelectedItem.ToString();
                if (txt_comentario.Text != null)
                    con.comentario = txt_comentario.Text;
               
                con.Save(nuevo_registro);
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

        private void FrmContratosDetalle_Edit_Load(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void LoadData()
        {
            if(con.nif != null)
            {
                txt_cups20.Text = con.cups20;
                txt_fecha_inicio.Text = con.fecha_inicio.ToString();
                txt_fecha_fin.Text = con.fecha_fin > DateTime.MinValue ? con.fecha_fin.ToString() : "";
                txt_qd.Text = con.qd.ToString("##.##");
                cmb_tarifa.Text = con.tarifa.ToString();
                cmb_tipo.Text = con.tipo.ToString();
                //if (con.tipo.ToString() == "INDEFINIDO")
                //{
                //    txt_fecha_fin.Text = "";
                //    //txt_fecha_fin.Enabled = false;
                //}                    
                //else
                //{
                //    txt_fecha_fin.Text = con.fecha_fin.ToString();
                //}
                    
            }
        }

        private void cmb_tipo_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if(cmb_tipo.Text == "INDEFINIDO")
            //{
            //    txt_fecha_fin.Text = "";
            //    //txt_fecha_fin.Enabled = false;
            //}else
            //{
            //    txt_fecha_fin.Enabled = true;
            //}


            
        }

        private void txt_qd_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
