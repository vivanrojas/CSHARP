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
    public partial class FrmPPAs : Form
    {
        EndesaBusiness.medida.PPAs ppas;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmPPAs()
        {
            usage.Start("Medida", "FrmPPAs", "N/A");
            InitializeComponent();
            ppas = new EndesaBusiness.medida.PPAs();
        }

        private void configuraciónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Configuración PPA´s";
            p.tabla = "ppas_param";
            p.esquemaString = "MED";
            p.Show();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_Exportar_Click(object sender, EventArgs e)
        {
            bool hayError = false;

            try
            {
                
                
                hayError = ppas.ExportarTabla_a_Excel_Agrupado();
                if(!hayError)
                    hayError = ppas.SubidaFTP();

                if (!hayError)
                {

                    this.lblUltimaExportacion.Text = string.Format("Última Exportación: {0}", 
                        ppas.UltimaExportacion().ToString("dd/MM/yyyy HH:mm:ss"));

                    MessageBox.Show("Exportación finalizada",
                        "Exportación PPAs",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                    
                }                    
                else
                    MessageBox.Show("Exportación finalizada con error",
                        "Exportación PPAs",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch(Exception ee)
            {
                MessageBox.Show(ee.Message,
                       "SubidaFTP",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }

        }

        private void FrmPPAs_Load(object sender, EventArgs e)
        {
            this.lblUltimaExportacion.Text = string.Format("Última Exportación: {0}",
                        ppas.UltimaExportacion().ToString("dd/MM/yyyy HH:mm:ss"));
        }

        private void acercaDePPAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmAyuda f = new FrmAyuda();
            f.rtb_Texto.Text = ppas.Ayuda();
            f.ShowDialog();
        }

        private void FrmPPAs_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Medida", "FrmPPAs", "N/A");
        }
    }
}
