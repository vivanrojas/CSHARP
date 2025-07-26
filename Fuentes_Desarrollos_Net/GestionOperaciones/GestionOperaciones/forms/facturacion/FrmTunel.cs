using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmTunel : Form
    {
        EndesaBusiness.utilidades.Param p = new EndesaBusiness.utilidades.Param("tunel_param", EndesaBusiness.servidores.MySQLDB.Esquemas.FAC);

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmTunel()
        {

            usage.Start("Facturación", "FrmTunel" ,"N/A");
            InitializeComponent();
        }

        private void btn_Carga_Excel_Click(object sender, EventArgs e)
        {
            
            EndesaBusiness.facturacion.Tunel t = new EndesaBusiness.facturacion.Tunel();
            DialogResult result = DialogResult.Yes;

            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Seleccione archivos Excel de Tunel";
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                result = MessageBox.Show("¿Desea procesar los Excels seleccionados?",
                  "Importación ficheros Excel Tunel",
                  MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    Cursor.Current = Cursors.WaitCursor;                    
                    
                    foreach (string fileName in d.FileNames)
                    {
                        t.CargaExcel(fileName, chk_Mensual.Checked);
                    }
                    Cursor.Current = Cursors.Default;

                    MessageBox.Show("Proceso Finalizado."
                        + System.Environment.NewLine
                        + System. Environment.NewLine
                        + "Puede mirar el resultado en la ruta:"
                        + System.Environment.NewLine
                        + p.GetValue("salida_excels", DateTime.Now,DateTime.Now),
                  "Importación ficheros Excel Tunel",
                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void acercaDeClausulaTúnelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmAyuda f = new FrmAyuda();
            f.rtb_Texto.Text = p.GetValue("ayuda", DateTime.Now, DateTime.Now);
            f.ShowDialog();
        }
        
        private void FrmTunel_Load(object sender, EventArgs e)
        {

        }

        private void FrmTunel_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmTunel" ,"N/A");
        }
    }
}
