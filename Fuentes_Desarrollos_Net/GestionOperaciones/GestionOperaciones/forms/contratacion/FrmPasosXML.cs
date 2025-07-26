using EndesaBusiness.factoring;
using OfficeOpenXml.LoadFunctions.Params;
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
    public partial class FrmPasosXML : Form
    {

        EndesaBusiness.xml.Extrasistemas extrasistemas;
        public FrmPasosXML()
        {
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void importarExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.xml.Extrasistemas t = new EndesaBusiness.xml.Extrasistemas();
            DialogResult result = DialogResult.Yes;

            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Seleccione la plantilla de Extrasistemas para generar XML";
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = false;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                foreach (string fileName in d.FileNames)
                {
                    t.CargaExcel(fileName);
                }
                Cursor.Current = Cursors.Default;
                LoadData();


            }
        }

        private void FrmPasosXML_Load(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            extrasistemas = new EndesaBusiness.xml.Extrasistemas();
            LoadData();
            Cursor.Current = Cursors.Default;
        }

        private void LoadData()
        {
            bool firstOnly = true;
            dgv.AutoGenerateColumns = false;
            dgv.DataSource = extrasistemas.Totales_Registros();


            txtInfo.Clear();
            foreach (string p in extrasistemas.lista_log)
            {
                if (firstOnly)
                {
                    txtInfo.AppendText(p);
                    firstOnly = false;
                }
                else
                {
                    txtInfo.AppendText(System.Environment.NewLine);
                    txtInfo.AppendText(p);
                }
                
            }
                
        }

        private void btn_importar_excel_Click(object sender, EventArgs e)
        {
            
            DialogResult result = DialogResult.Yes;

            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Seleccione la plantilla de Extrasistemas para generar XML";
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = false;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                foreach (string fileName in d.FileNames)
                {
                    extrasistemas.CargaExcel(fileName);
                }
                Cursor.Current = Cursors.Default;
                LoadData();

                //MessageBox.Show("Proceso Finalizado."
                //        + System.Environment.NewLine
                //        + System.Environment.NewLine
                //        //+ "Puede mirar el resultado en la ruta:"
                //        + System.Environment.NewLine,
                //  //+ p.GetValue("salida_excels", DateTime.Now, DateTime.Now),
                //  "Proceso XML Extrasistemas",
                //  MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            extrasistemas.Crea_XML();
            LoadData();
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
    }
}
