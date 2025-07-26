using EndesaBusiness.servidores;
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
    public partial class FrmInventarioBI : Form
    {
        EndesaBusiness.utilidades.Param param;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmInventarioBI()
        {
            usage.Start("Medida", "FrmInventarioBI", "N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmInventarioBI_Load(object sender, EventArgs e)
        {
            param = new EndesaBusiness.utilidades.Param("t_ed_h_gest_diar_ps_b2b_param", MySQLDB.Esquemas.MED);
            this.lbl_ultima_actualización.Text = string.Format("Última fecha de ejecución del proceso: {0}", 
                Convert.ToDateTime(param.GetValue("ultima_actualizacion")).ToString("dd/MM/yyyy HH:mm:ss"));

           btn_relanzar_proceso.Enabled = param.GetValue("estado_proceso") == "ejecutado";

        }

        private void btn_relanzar_proceso_Click(object sender, EventArgs e)
        {
            param.UpdateParameter("estado_proceso", "lanzar");

            MessageBox.Show("Se ha puesto el proceso de actualización de inventario BI en ejecución."
                + System.Environment.NewLine
                + "Se le avisará con un email cuando finalize el proceso."
                , "Actualización Programada",
                   MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            btn_relanzar_proceso.Enabled = false;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void lbl_ultima_actualización_Click(object sender, EventArgs e)
        {

        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "t_ed_h_gest_diar_ps_b2b_param";
            p.esquemaString = "MED";
            p.cabecera = "Parámetros Inventario BI";
            p.Show();
        }

        private void FrmInventarioBI_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Medida", "FrmInventarioBI", "N/A");
        }
    }
}
