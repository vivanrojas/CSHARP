using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.contratacion
{
    public partial class FrmComentariosEdit : Form
    {
        public EndesaEntity.contratacion.xxi.Comentario comentario { get; set; }
        public EndesaBusiness.contratacion.eexxi.Comentarios.EditStatus editStatus { get; set; }
        
        public FrmComentariosEdit()
        {
            InitializeComponent();
        }
                

        private void FrmComentariosEdit_Load(object sender, EventArgs e)
        {
            
            
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void TxtComentario_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (txtComentario.Text == "")
                errorProvider.SetError(txtComentario, "Debe informar el comentario");
            else
            {
                EndesaBusiness.contratacion.eexxi.Comentarios comment
                    = new EndesaBusiness.contratacion.eexxi.Comentarios(comentario.id_comentario);

                comment.id_comentario = comentario.id_comentario;
                
                
                switch (editStatus)
                {
                    case EndesaBusiness.contratacion.eexxi.Comentarios.EditStatus.NEW:
                        comment.linea = comment.UltimaLinea(comentario.id_comentario) + 1;
                        comment.comentario = txtComentario.Text;
                        comment.Add();
                        break;
                    case EndesaBusiness.contratacion.eexxi.Comentarios.EditStatus.EDIT:
                        comment.linea = comentario.linea;
                        comment.comentario = txtComentario.Text;
                        comment.Update();
                        break;
                }
                
                this.Close();
            }
                

        }
    }
}
