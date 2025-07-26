using EndesaBusiness.contratacion;
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
    public partial class FrmPS_AT_HIST : Form
    {
        EndesaBusiness.contratacion.PS_AT_HIST ps_at_hist;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        public FrmPS_AT_HIST()
        {
            usage.Start("Contratación", "FrmPS_AT_HIST" ,"N/A");
            ps_at_hist = new EndesaBusiness.contratacion.PS_AT_HIST();
            InitializeComponent();
            InitListBox();
            this.listBoxFechas.SelectedIndex = 0;
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmPS_AT_HIST_Load(object sender, EventArgs e)
        {

        }

        private void InitListBox()
        {
            for (int i = listBoxFechas.Items.Count - 1; i == 0; i--)
            {
                listBoxFechas.Items.RemoveAt(i);
            }

            List<DateTime> lista_fechas = ps_at_hist.ListaFechaAnexion();

            for (int i = 0; i < lista_fechas.Count; i++)
            {
                listBoxFechas.Items.Add(lista_fechas[i]);
            }
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {            
            
                SaveFileDialog save = new SaveFileDialog();
                save.Title = "Ubicación del informe";
                save.AddExtension = true;
                save.DefaultExt = "xlsx";
                save.Filter = "Ficheros xslx (*.xlsx)|*.*";
                DialogResult result = save.ShowDialog();
                if (result == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    ps_at_hist.Exporta_PS_AT_Excel(Convert.ToDateTime(listBoxFechas.SelectedItem), save.FileName);
                    Cursor.Current = Cursors.Default;

                    DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (result2 == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(save.FileName);
                    }
                }
            
        }

        private void FrmPS_AT_HIST_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Contratación", "FrmPS_AT_HIST" ,"N/A");
        }
    }
}
