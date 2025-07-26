using EndesaEntity.medida;
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
using System.Xml.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace GestionOperaciones.forms.contratacion
{
    public partial class FrmPlantillaXML : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmPlantillaXML()
        {

            usage.Start("Contratación", "FrmPlantillaXML", "N/A");
            InitializeComponent();
        }

        private void btn_generar_xml_Click(object sender, EventArgs e)
        {
            EndesaBusiness.xml.XMLFunciones xml = new EndesaBusiness.xml.XMLFunciones();

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    Cursor.Current = Cursors.WaitCursor;
                    xml.CargaDatosExcel_XML_Platillas(txt_datos_excel.Text, txt_plantilla.Text, fbd.SelectedPath);
                    Cursor.Current = Cursors.Default;

                }


            }
        }

        private void btn_plantilla_Click(object sender, EventArgs e)
        {
            DialogResult result = DialogResult.Yes;
            EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301 xml;


            try
            {
                OpenFileDialog d = new OpenFileDialog();
                d.Title = "Carga XML";
                d.Filter = "xml files|*.xml";
                d.Multiselect = false;

                if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (string fileName in d.FileNames)
                    {
                        txt_plantilla.Text = fileName;                        
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                "Error en carga XML",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        private void btn_excel_datos_Click(object sender, EventArgs e)
        {
            DialogResult result = DialogResult.Yes;
            EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301 xml;


            try
            {
                OpenFileDialog d = new OpenFileDialog();
                d.Title = "Carga Excel";
                d.Filter = "xlsx files|*.xlsx;*.xls";
                d.Multiselect = false;

                if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (string fileName in d.FileNames)
                    {
                        txt_datos_excel.Text = fileName;

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                "Error en carga XML",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        private void FrmPlantillaXML_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Contratación", "FrmPlantillaXML", "N/A");
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
