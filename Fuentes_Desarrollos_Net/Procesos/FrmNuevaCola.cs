using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Procesos
{
    public partial class FrmNuevaCola : Form
    {
        business.Cola cola;

        public FrmNuevaCola()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {

            if (cmb_cola.Text == "")
            {
                if (txt_nombreCola.Text == null || txt_nombreCola.Text == "")
                    errorProvider.SetError(txt_nombreCola, "Debe informar un valor");
                else if (txt_descripcion.Text == null || txt_descripcion.Text == "")
                    errorProvider.SetError(txt_descripcion, "Debe informar un valor");
                else if (txt_mail.Text == null || txt_mail.Text == "")
                    errorProvider.SetError(txt_mail, "Debe informar un valor");
                else
                {
                    cola.cola = txt_nombreCola.Text;
                    cola.descripcion = txt_descripcion.Text;
                    cola.mail_aviso = txt_mail.Text;
                    cola.Save();                    
                    using (Main m = new Main(txt_nombreCola.Text, false))
                    {
                        m.ShowDialog();
                    }
                    

                }
            }
            else
            {
                Main m = new Main(cmb_cola.Text, false);
                m.Text = this.Text = "Cola de Procesos de " + cmb_cola.Text;
                m.Show();                

            }
            

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmNuevaCola_Load(object sender, EventArgs e)
        {
            cola = new business.Cola();
            CargaCombo();
        }

        private void CargaCombo()
        {
            List<string> lista = cola.dic.Select(z => z.Key).ToList();
            for (int i = 0; i < lista.Count; i++)
                cmb_cola.Items.Add(lista[i]);
        }
       

       
    }
}
