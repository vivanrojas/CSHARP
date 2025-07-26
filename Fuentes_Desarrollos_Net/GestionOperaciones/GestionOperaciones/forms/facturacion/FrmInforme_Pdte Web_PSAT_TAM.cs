using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmInforme_Pdte_Web_PSAT_TAM : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmInforme_Pdte_Web_PSAT_TAM()
        {

            usage.Start("Facturación", "FrmInforme_Pdte_Web_PSAT_TAM" ,"N/A");
            InitializeComponent();
        }

        private void btn_excel_Click(object sender, EventArgs e)
        {
            FileInfo file;
            SaveFileDialog save;            
            string ruta_salida_archivo = "";
            EndesaBusiness.medida.Pendiente pendiente = new EndesaBusiness.medida.Pendiente();

            save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                file = new FileInfo(save.FileName);
                if (file.Exists)
                    file.Delete();

                ruta_salida_archivo = save.FileName;
                Cursor = Cursors.WaitCursor;
                pendiente.InformePendWeb_PSAT_TAM_V2(ruta_salida_archivo, false);
                Cursor = Cursors.Default;

                DialogResult result2 = MessageBox.Show("Infome Terminado." +
                       System.Environment.NewLine +
                       "¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(file.FullName);
                }

            }

            
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "global_param";
            p.esquemaString = "AUX";
            p.cabecera = "Parámetros Pdte Web + PSAT + TAM";
            p.Show();
        }

        private void FrmInforme_Pdte_Web_PSAT_TAM_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmInforme_Pdte_Web_PSAT_TAM" ,"N/A");
        }
    }
}
