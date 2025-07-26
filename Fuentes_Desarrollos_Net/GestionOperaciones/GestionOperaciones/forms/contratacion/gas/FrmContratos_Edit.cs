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
    
    public partial class FrmContratos_Edit : Form
    {
        EndesaBusiness.contratacion.gestionATRGas.Distribuidoras dis;
        public EndesaBusiness.contratacion.gestionATRGas.ContratosGas con { get; set; }
        
        public FrmContratos_Edit()
        {
            InitializeComponent();
            dis = new EndesaBusiness.contratacion.gestionATRGas.Distribuidoras(true);
        }

        

        private void btnOK_Click(object sender, EventArgs e)
        {

            errorProvider.Clear();

            if (txt_cnifdnic.Text == null || txt_cnifdnic.Text == "")
                errorProvider.SetError(txt_cnifdnic, "El campo debe estar informado.");
            else if (txt_dapersoc.Text == null || txt_dapersoc.Text == "")
                errorProvider.SetError(txt_dapersoc, "El campo debe estar informado.");
            else if (cmb_distribuidora.SelectedItem == null)
                errorProvider.SetError(cmb_distribuidora, "Debe seleccionar una distribuidora válida.");
            else if (txt_cups20.Text == null || txt_cups20.Text == "")
                errorProvider.SetError(txt_cups20, "El campo debe estar informado.");
            else if (txt_cups20.Text.Length != 20)
                errorProvider.SetError(txt_cups20, "El CUPS  no tiene 20 caracteres.");
            else if (cmb_tramitacion.SelectedItem == null)
                errorProvider.SetError(cmb_tramitacion, "Debe seleccionar una tramitación válida.");
            else
            {

                con.cnifdnic = txt_cnifdnic.Text;
                con.dapersoc = txt_dapersoc.Text;
                con.distribuidora = cmb_distribuidora.SelectedItem.ToString(); //.Text;
                con.cups20 = txt_cups20.Text;
                con.comentarios_descuadres = txt_comentarios_descuadres.Text;
                con.comentarios_contratacion = txt_comentarios_contratacion.Text;
                con.tramitacion = cmb_tramitacion.SelectedItem.ToString();
                con.Save();
                this.Close();
            }
                
            
            
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
               

        
        private void FrmContratos_Edit_Load(object sender, EventArgs e)
        {
            for (int i = cmb_distribuidora.Items.Count - 1; i == 0; i--)
            {
                cmb_distribuidora.Items.RemoveAt(i);
            }


            foreach (KeyValuePair<string,EndesaEntity.Table_atrgas_distribuidoras> p in dis.l_distribuidoras)
            {
                cmb_distribuidora.Items.Add(p.Key);
            }

            
            
                
            
        }
    }
}
