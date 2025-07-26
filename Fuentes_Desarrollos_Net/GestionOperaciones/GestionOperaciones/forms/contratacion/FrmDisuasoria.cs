using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;

namespace GestionOperaciones.forms.contratacion
{
    public partial class FrmDisuasoria : Form
    {

        //List<GO.contratacion.Diasuario> lista_cups;
        //List<GO.contratacion.Disuasorio_Resumen> lista_cups_resumen;

        EndesaBusiness.contratacion.TarifaDisuasoria tarifaDisuasoria;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmDisuasoria()
        {

            usage.Start("Contratación", "FrmDisuasoria" ,"N/A");

            InitializeComponent();
            tarifaDisuasoria = new EndesaBusiness.contratacion.TarifaDisuasoria();

            this.txtMaxFecha.Text = tarifaDisuasoria.UltimaImportacion();

            if(Convert.ToDateTime(this.txtMaxFecha.Text).Month < DateTime.Now.Month)
                MessageBox.Show("La extracción C70CCNAE no está actualizada."
                    + System.Environment.NewLine
                    + "Por favor, pida una nueva extracción a sistemas antes de enviar el informe.", 
                    "Actualizar Extracción",
                   MessageBoxButtons.OK, MessageBoxIcon.Information);

            //this.txtFecha.Text = tarifaDisuasoria.UltimaFechaPS_Historico();
            //this.cmbPotencia.SelectedIndex = 0;
            //CargadgvResumen(this.cmbPotencia.SelectedIndex == 0);
            CargadgvResumen();
            //CargardgvCUPS(this.cmbPotencia.SelectedIndex == 0);
        }

       



        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        //private void CargardgvCUPS(bool mayor50)
        //{
        //    try
        //    {
        //        Cursor.Current = Cursors.WaitCursor;
        //        dgvCups.DataSource = tarifaDisuasoria.CargardgvCUPS(mayor50, txtFecha.Value);
        //        Cursor.Current = Cursors.Default;
        //    }
        //    catch(Exception e)
        //    {
        //        MessageBox.Show(e.Message,
        //       "Error en la ejecución de la consulta C70CCNAE_CUPS",
        //       MessageBoxButtons.OK,
        //       MessageBoxIcon.Error);
        //    }
        //}



        //private void CargadgvResumen(bool mayor50)
        private void CargadgvResumen()
        {

           

            try
            {                
                Cursor.Current = Cursors.WaitCursor;
                tarifaDisuasoria.CargadgvResumen();
                dgvResumen.DataSource = tarifaDisuasoria.lista_Disusoria_informe;
                //dgvResumen.DataSource = tarifaDisuasoria.CargadgvResumen(mayor50, txtFecha.Value);
                Cursor.Current = Cursors.Default;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la ejecución de la consulta C70CCNAE_CUPS_RESUMEN",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

        }

        private void btnExcel_Click(object sender, EventArgs e)
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
                tarifaDisuasoria.ExportExcel(save.FileName);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }

           
        }



        

        private void menuCopyPaste_Opening(object sender, CancelEventArgs e)
        {

        }

        private void importarToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }



        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void C70CCNAE_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "txt files|*.txt";
            d.Multiselect = true;
            EndesaBusiness.contratacion.C70CCNAE fichero;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fichero = new EndesaBusiness.contratacion.C70CCNAE();

                foreach (string fileName in d.FileNames)
                {
                    //fichero = new EndesaBusiness.contratacion.C70CCNAE(fileName);
                    fichero.Carga_C70CCNAE(fileName);
                }

                MessageBox.Show("La importación ha concluido correctamente.",
                    "Importación ficheros C70CCNAE",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                //this.CargardgvCUPS(dtFFACTDES.Value.ToString("yyyy-MM-dd"), dtFFACTHAS.Value.ToString("yyyy-MM-dd"), txtcupsree.Text, txtlote.Text);
                //adif.Adif_fichero_factura ff = new adif.Adif_fichero_factura();
            }
        }



        private void salirToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Close();
        }

        private void cmbPotencia_SelectedIndexChanged(object sender, EventArgs e)
        {
            //CargadgvResumen(this.cmbPotencia.SelectedIndex == 0);
            CargadgvResumen();
            //CargardgvCUPS(this.cmbPotencia.SelectedIndex == 0);

        }

        private void cNAEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.FrmCNAE f = new forms.FrmCNAE();
            //f.Show();
        }

        private void excComplToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.FrmDisuasoria_Param f = new forms.FrmDisuasoria_Param();
            //f.Show();
        }

        private void FrmDisuasoria_Load(object sender, EventArgs e)
        {

        }

        private void CmdSearch_Click(object sender, EventArgs e)
        {
            //CargadgvResumen(this.cmbPotencia.SelectedIndex == 0);
            //CargardgvCUPS(this.cmbPotencia.SelectedIndex == 0);
        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "tarifa_disuasoria_param";
            p.esquemaString = "CON";
            p.cabecera = "Parámetros Tarifa Disuasoria";
            p.Show();
        }

        private void FrmDisuasoria_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Contratación", "FrmDisuasoria" ,"N/A");
        }
    }
}
