using EndesaEntity;
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
    public partial class FrmAdifIventario_Edit : Form
    {
        EndesaBusiness.adif.Provincias provincias;
        private const char SignoDecimal = ',';

        public EndesaBusiness.adif.InventarioFunciones inventario { get; set; }

        public FrmAdifIventario_Edit()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (txt_cups20.Text == null || txt_cups20.Text == "")
                errorProvider.SetError(txt_cups20, "El campo debe estar informado.");
            else if (txt_lote.Text == null || txt_lote.Text == "")
                errorProvider.SetError(txt_lote, "El campo debe estar informado.");
            else if (txt_fd.Value == null || txt_fd.Value == null)
                errorProvider.SetError(txt_fd, "El campo debe estar informado.");
            else if (txt_fh.Value == null || txt_fh.Value == null)
                errorProvider.SetError(txt_fd, "El campo debe estar informado.");
            else if (txt_zona.Text == null || txt_zona.Text == "")
                errorProvider.SetError(txt_zona, "El campo debe estar informado.");
            else if (txt_codigo.Text == null || txt_codigo.Text == "")
                errorProvider.SetError(txt_codigo, "El campo debe estar informado.");
            else if (txt_nombre_suministro.Text == null || txt_nombre_suministro.Text == "")
                errorProvider.SetError(txt_nombre_suministro, "El campo debe estar informado.");
            else if (txt_distribuidora.Text == null || txt_distribuidora.Text == "")
                errorProvider.SetError(txt_distribuidora, "El campo debe estar informado.");
            else if (txt_tarifa.Text == null || txt_tarifa.Text == "")
                errorProvider.SetError(txt_tarifa, "El campo debe estar informado.");
            
            else
            {
                inventario.cups20 = txt_cups20.Text;
                inventario.ffactdes = txt_fd.Value;
                inventario.ffacthas = txt_fh.Value;
                inventario.lote = Convert.ToInt32(txt_lote.Text);
                inventario.zona = txt_zona.Text;
                inventario.codigo = txt_codigo.Text;
                inventario.nombre_punto_suministro = txt_nombre_suministro.Text;
                inventario.distribuidora = txt_distribuidora.Text;
                inventario.tarifa = txt_tarifa.Text;
                inventario.devolucion_de_energia = chk_devolucion_energia.Checked;
                inventario.medida_en_baja = chk_medida_en_baja.Checked;
                inventario.comentarios = txt_comentarios.Text;

                inventario.provincia = cmb_provincia.Text;
                inventario.comunidad_autonoma = txt_comunidad_autonoma.Text;

                if(txt_sistema_traccion.Text != "")
                    inventario.sitema_traccion = txt_sistema_traccion.Text;

                if(txt_grupo.Text != "")
                    inventario.grupo = txt_grupo.Text;

                if (txt_valor_kvas.Text != "")
                    inventario.valor_kvas = Convert.ToDouble(txt_valor_kvas.Text);

                if (txt_perdidas.Text != "")
                    inventario.perdidas = Convert.ToDouble(txt_perdidas.Text);

                if (txt_multipuntos.Text != "")
                    inventario.multipunto_num_principales = Convert.ToInt32(txt_multipuntos.Text);               
                
                
                if (txt_p1.Text != "")
                    inventario.p[0] = Convert.ToInt32(txt_p1.Text);

                if (txt_p2.Text != "")
                    inventario.p[1] = Convert.ToInt32(txt_p2.Text);

                if (txt_p3.Text != "")
                    inventario.p[2] = Convert.ToInt32(txt_p3.Text);

                if (txt_p4.Text != "")
                    inventario.p[3] = Convert.ToInt32(txt_p4.Text);

                if (txt_p5.Text != "")
                    inventario.p[4] = Convert.ToInt32(txt_p5.Text);

                if (txt_p6.Text != "")
                    inventario.p[5] = Convert.ToInt32(txt_p6.Text);

                inventario.Save();
                this.Close();

            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void FrmAdifIventario_Edit_Load(object sender, EventArgs e)
        {
            InitComboBox();
        }


        private void InitComboBox()
        {
            provincias = new EndesaBusiness.adif.Provincias();

            for (int i = cmb_provincia.Items.Count - 1; i >= 0; i--)
                cmb_provincia.Items.RemoveAt(i);

            foreach (KeyValuePair<string, string> p in provincias.dic_provincias)
                cmb_provincia.Items.Add(p.Key);

        }

        private void txt_valor_kvas_KeyPress(object sender, KeyPressEventArgs e)
        {
            var textBox = (TextBox)sender;            
            e.Handled = !char.IsDigit(e.KeyChar) 
                        && !char.IsControl(e.KeyChar)
                        && (e.KeyChar != SignoDecimal
                            || textBox.SelectionStart == 0
                            || textBox.Text.Contains(SignoDecimal));
        }

        private void txt_perdidas_KeyPress(object sender, KeyPressEventArgs e)
        {
            var textBox = (TextBox)sender;
            e.Handled = !char.IsDigit(e.KeyChar)
                        && !char.IsControl(e.KeyChar)
                        && (e.KeyChar != SignoDecimal
                            || textBox.SelectionStart == 0
                            || textBox.Text.Contains(SignoDecimal));
        }

        private void txt_multipuntos_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_multipuntos_KeyPress(object sender, KeyPressEventArgs e)
        {
            var textBox = (TextBox)sender;
            e.Handled = !char.IsDigit(e.KeyChar)
                        && !char.IsControl(e.KeyChar)
                        && (e.KeyChar != SignoDecimal
                            || textBox.SelectionStart == 0);
        }

        private void txt_p1_KeyPress(object sender, KeyPressEventArgs e)
        {
            var textBox = (TextBox)sender;
            e.Handled = !char.IsDigit(e.KeyChar)
                        && !char.IsControl(e.KeyChar)
                        && (e.KeyChar != SignoDecimal
                            || textBox.SelectionStart == 0
                            || textBox.Text.Contains(SignoDecimal));
        }

        private void txt_p2_KeyPress(object sender, KeyPressEventArgs e)
        {
            var textBox = (TextBox)sender;
            e.Handled = !char.IsDigit(e.KeyChar)
                        && !char.IsControl(e.KeyChar)
                        && (e.KeyChar != SignoDecimal
                            || textBox.SelectionStart == 0
                            || textBox.Text.Contains(SignoDecimal));
        }

        private void txt_p3_KeyPress(object sender, KeyPressEventArgs e)
        {
            var textBox = (TextBox)sender;
            e.Handled = !char.IsDigit(e.KeyChar)
                        && !char.IsControl(e.KeyChar)
                        && (e.KeyChar != SignoDecimal
                            || textBox.SelectionStart == 0
                            || textBox.Text.Contains(SignoDecimal));
        }

        private void txt_p4_KeyPress(object sender, KeyPressEventArgs e)
        {
            var textBox = (TextBox)sender;
            e.Handled = !char.IsDigit(e.KeyChar)
                        && !char.IsControl(e.KeyChar)
                        && (e.KeyChar != SignoDecimal
                            || textBox.SelectionStart == 0
                            || textBox.Text.Contains(SignoDecimal));
        }

        private void txt_p5_KeyPress(object sender, KeyPressEventArgs e)
        {
            var textBox = (TextBox)sender;
            e.Handled = !char.IsDigit(e.KeyChar)
                        && !char.IsControl(e.KeyChar)
                        && (e.KeyChar != SignoDecimal
                            || textBox.SelectionStart == 0
                            || textBox.Text.Contains(SignoDecimal));
        }

        private void txt_p6_KeyPress(object sender, KeyPressEventArgs e)
        {
            var textBox = (TextBox)sender;
            e.Handled = !char.IsDigit(e.KeyChar)
                        && !char.IsControl(e.KeyChar)
                        && (e.KeyChar != SignoDecimal
                            || textBox.SelectionStart == 0
                            || textBox.Text.Contains(SignoDecimal));
        }
    }
}
