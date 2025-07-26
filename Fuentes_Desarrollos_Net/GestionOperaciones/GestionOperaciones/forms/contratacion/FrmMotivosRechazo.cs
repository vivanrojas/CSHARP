using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.contratacion
{
    public partial class FrmMotivosRechazo : Form
    {
        EndesaBusiness.contratacion.MotivosRechazos mot;
        EndesaBusiness.contratacion.TiposMotivosRechazo tipo_mot;
        public FrmMotivosRechazo()
        {
            InitializeComponent();
        }

        

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmMotivosRechazo_Load(object sender, EventArgs e)
        {
            mot = new EndesaBusiness.contratacion.MotivosRechazos();
            tipo_mot = new EndesaBusiness.contratacion.TiposMotivosRechazo();

            cmb_Motivos.Items.Add("");
            for (int i = 0; i < tipo_mot.list.Count; i++)
                cmb_Motivos.Items.Add(tipo_mot.list[i].motivo);

            cmb_Motivos.DropDownStyle = ComboBoxStyle.DropDownList;

            LoadData();
            
        }


        private void LoadData()
        {
            
            List<EndesaEntity.contratacion.MotivosRechazo> lista = mot.dic.Values.ToList();           

            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;
            lbl_total_registros.Text = string.Format("Nº de Rechazos: {0:#,##0}", lista.Count());
            LoadDataDetail();
        }

        private void LoadDataDetail()
        {
            EndesaEntity.contratacion.MotivosRechazo c = mot.dic.First().Value;
            if (c.cups != null)
                txt_CUPS.Text = c.cups;
            if (c.clienteActualizado != null)
                txt_Cliente_Actualizado.Text = c.clienteActualizado;
            if (c.numSolAtr != 0)
                txt_NumSolAtr.Text = c.numSolAtr.ToString();
            if (c.fechaRechazoSol > DateTime.MinValue)
                txt_FecRechazoSol.Text = c.fechaRechazoSol.Date.ToString("dd/MM/yyyy");
            if (c.tipoSolicitud != null)
                txt_TipoSol.Text = c.tipoSolicitud;
            if (c.rechazoPdte != null)
                txt_RechazoPdte.Text = c.rechazoPdte.ToString();
            if (c.motivos != null && c.motivos != "")
                cmb_Motivos.Text = c.motivos;
            else
                cmb_Motivos.Text = "";
            if (c.comentario != null && c.comentario != "")
                rtxt_Observaciones.Text = c.comentario;
            else
                rtxt_Observaciones.Text = "";

            btnCancel.Enabled = false;
            btnOK.Enabled = false;

        }


        private void groupBox2_Enter(object sender, EventArgs e)
        {
           
        }


        private void Datos_dgv(DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewCell cups = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[0];

                DataGridViewCell clienteActualizado = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[1];

                DataGridViewCell numSolAtr = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[2];

                DataGridViewCell fecRechazoSol = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[3];

                DataGridViewCell tipoSolicitud = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[4];

                DataGridViewCell rechazoPdte = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[5];

                DataGridViewCell rechazmotivos = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[6];

                DataGridViewCell comentarios = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[7];

                if (cups.Value != null)
                    txt_CUPS.Text = cups.Value.ToString();
                if (clienteActualizado.Value != null)
                    txt_Cliente_Actualizado.Text = clienteActualizado.Value.ToString();
                if (numSolAtr.Value != null)
                    txt_NumSolAtr.Text = numSolAtr.Value.ToString();
                if (fecRechazoSol.Value != null)
                    txt_FecRechazoSol.Text = Convert.ToDateTime(fecRechazoSol.Value).Date.ToString("dd/MM/yyyy");
                if (tipoSolicitud.Value != null)
                    txt_TipoSol.Text = tipoSolicitud.Value.ToString();
                if (rechazoPdte.Value != null)
                    txt_RechazoPdte.Text = rechazoPdte.Value.ToString();
                if (rechazmotivos.Value != null)
                    cmb_Motivos.Text = rechazmotivos.Value.ToString();
                else
                    cmb_Motivos.Text = null;
                if (comentarios.Value != null)
                    rtxt_Observaciones.Text = comentarios.Value.ToString();
                else
                    rtxt_Observaciones.Text = null;

                            
            }
        }

        private void dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Datos_dgv(e);
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if ((cmb_Motivos.Text == null)
                && (rtxt_Observaciones.Text == null || rtxt_Observaciones.Text == ""))
            {
                errorProvider.SetError(cmb_Motivos, "El campo debe estar informado.");
                errorProvider.SetError(rtxt_Observaciones, "El campo debe estar informado.");

            }
            else
            {
                DialogResult result_1 = MessageBox.Show("¿Desea guardar los cambios?",
               "Motivos Rechazo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result_1 == DialogResult.Yes)
                {
                    mot.numSolAtr = Convert.ToInt64(txt_NumSolAtr.Text);
                    mot.motivos = cmb_Motivos.Text;
                    if(rtxt_Observaciones.Text != null)                      
                        mot.comentario = rtxt_Observaciones.Text;
                    mot.Save();
                    LoadData();
                }

                    
            }
        }

        private void dgv_KeyUp(object sender, KeyEventArgs e)
        {
            //btnOK.Enabled = false;
            //btnCancel.Enabled = false;
            //Datos_dgv(e);
        }

        private void dgv_KeyDown(object sender, KeyEventArgs e)
        {
            //btnOK.Enabled = false;
            //btnCancel.Enabled = false;
            //Datos_dgv(e);
        }

        private void cmb_Motivos_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnCancel.Enabled = true;
            btnOK.Enabled = true;
        }

        private void rtxt_Observaciones_TextChanged(object sender, EventArgs e)
        {
            btnCancel.Enabled = true;
            btnOK.Enabled = true;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void btnOK_Click_1(object sender, EventArgs e)
        {
            if ((cmb_Motivos.Text == null)
               && (rtxt_Observaciones.Text == null || rtxt_Observaciones.Text == ""))
            {
                errorProvider.SetError(cmb_Motivos, "El campo debe estar informado.");
                errorProvider.SetError(rtxt_Observaciones, "El campo debe estar informado.");

            }
            else
            {
                DialogResult result_1 = MessageBox.Show("¿Desea guardar los cambios?",
               "Motivos Rechazo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result_1 == DialogResult.Yes)
                {
                    mot.numSolAtr = Convert.ToInt64(txt_NumSolAtr.Text);
                    mot.motivos = cmb_Motivos.Text;
                    if (rtxt_Observaciones.Text != null)
                        mot.comentario = rtxt_Observaciones.Text;
                    mot.Save();
                    LoadData();
                }


            }
        }
    }
}
