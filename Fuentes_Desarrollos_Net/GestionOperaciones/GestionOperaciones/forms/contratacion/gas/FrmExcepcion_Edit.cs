using EndesaBusiness.contratacion.gestionATRGas;
using EndesaEntity.contratacion;
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
    public partial class FrmExcepcion_Edit : Form
    {
        public bool nuevo_registro { get; set; }
        public Int32 id_excepcion {  get; set; }
        public FrmExcepcion_Edit()
        {
            InitializeComponent();
        }

        private void btn_excepcion_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_excepcion_OK_Click(object sender, EventArgs e)
        {
            bool todoOK = true;

            //Falta añadir control fechas coherentes y que no solapen con otras excepciones para una misma distribuidora 
            if (Convert.ToDateTime(dateTimePicker_fd.Text)> Convert.ToDateTime(dateTimePicker_fh.Text))
            {
                errorProvider_frmExcepcion.SetError(dateTimePicker_fd, "La fecha desde no puede ser mayor que la fecha hasta.");
                todoOK = false;
            }

            if (todoOK)
            {
                if (this.nuevo_registro)
                {
                    DialogResult result = MessageBox.Show("¿Está seguro que desea añadir la nueva excepción?", "Añadir excepción",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        EndesaBusiness.contratacion.gestionATRGas.Excepciones.Add(cmb_lista_distribuidoras.Text, Convert.ToDateTime(dateTimePicker_fd.Text), Convert.ToDateTime(dateTimePicker_fh.Text));
                        MessageBox.Show("La excepción en tramitación para el grupo " + cmb_lista_distribuidoras.Text + " se ha añadido correctamente",
                                    "Excepción añadida",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                        this.Close();
                    }

                }
                else
                {
                    DialogResult result = MessageBox.Show("¿Está seguro que desea modificar la excepción?", "Modificar excepción",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        EndesaBusiness.contratacion.gestionATRGas.Excepciones.Update(id_excepcion, cmb_lista_distribuidoras.Text, Convert.ToDateTime(dateTimePicker_fd.Text), Convert.ToDateTime(dateTimePicker_fh.Text));
                        MessageBox.Show("La excepción en tramitación para el grupo " + cmb_lista_distribuidoras.Text + " se ha modificado correctamente",
                            "Excepción modificada",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        this.Close();
                    }
                }
            }
            
        }
    }
}
