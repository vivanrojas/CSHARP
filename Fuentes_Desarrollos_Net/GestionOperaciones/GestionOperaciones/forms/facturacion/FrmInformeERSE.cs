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
using System.Xml;
using System.Globalization;
using System.Diagnostics.Eventing.Reader;
using MS.WindowsAPICodePack.Internal;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmInformeERSE : Form
    {

        DataTable dt;
        EndesaBusiness.facturacion.InformeErse erse;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmInformeERSE()
        {

            DateTime fd;
            DateTime fh;
            DateTime mesAnterior;
            int anio;
            int mes;
            int dias_del_mes;


            usage.Start("Facturación", "FrmInformeERSE" ,"N/A");

            InitializeComponent();

            //this.cmbTipoFecha.SelectedIndex = 0;
            //this.cmbEntorno.SelectedIndex = 0;
            //this.cmbSistema.SelectedIndex = 0;

            mesAnterior = DateTime.Now.AddMonths(-1);
            anio = mesAnterior.Year;
            mes = mesAnterior.Month;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            fh = new DateTime(anio, mes, dias_del_mes);
            fd = new DateTime(fh.Year, fh.Month, 1);            

            this.txtFFACTURADES.Value = fd;
            this.txtFFACTURAHAS.Value = fh;           

        }

        private void Carga_adgv_BTE()
        {

            try
            {
                
                //dgvBTE.AutoGenerateColumns = true;
                //dgvBTE.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
                //dgvBTE.RowHeadersVisible = false;
                //dgvBTE.DataSource = erse.fact_list_BTE.Select(z => z.Value).ToList();                

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("La consulta de BTE que ha realizado no devuelve ningún resultado.",
                "La consulta no devuelte datos.",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
                }
                
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                    "Carga_adgv_BTE",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void Carga_adgv_BTN()
        {

            try
            {

                //dgvBTN.AutoGenerateColumns = true;
                //dgvBTN.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
                //dgvBTN.RowHeadersVisible = false;
                //dgvBTN.DataSource = erse.fact_list_BTN.Select(z => z.Value).ToList();

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("La consulta de BTN que ha realizado no devuelve ningún resultado.",
                "La consulta no devuelte datos.",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "Carga_adgv_BTN",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }
        private void Carga_adgv_MT()
        {  
           
            
             
            try
            {
                //dgvMT.AutoGenerateColumns = true;
                //dgvMT.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
                //dgvMT.RowHeadersVisible = false;
                //dgvMT.DataSource = erse.fact_list_MT.Select(z => z.Value).ToList();

                //MySqlDataAdapter da = new MySqlDataAdapter(command);
                //dt = new DataTable();
                //da.Fill(dt);
                //dgvFacturas.AutoGenerateColumns = false;
                //dgvFacturas.DataSource = dt;



                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("La consulta de MT que ha realizado no devuelve ningún resultado.",
                "La consulta no devuelte datos.",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
                }

                

            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
                  "Carga_adgv_MT",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }


        private void FrmPSAT_Load(object sender, EventArgs e)
        {
            // this.lblUpdate.Text = string.Format("Última actualización: {0}", this.UltimaActualizacion());
            // this.Carga_adgv_MT();

            
        }
        

        private void acercaDeInformeERSEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmInformeERSE_Ayuda f = new forms.facturacion.FrmInformeERSE_Ayuda();
            f.ShowDialog();
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            //Están rellenos los campos correctamente?
            
            SaveFileDialog save;
            DialogResult result;
            erse = new EndesaBusiness.facturacion.InformeErse();
            erse.fd = Convert.ToDateTime(txtFFACTURADES.Text);
            erse.fh = Convert.ToDateTime(this.txtFFACTURAHAS.Text);
            //erse.usarFechaFactura = cmbTipoFecha.SelectedIndex == 0;

               
            // Sistema SAP 
            save = new SaveFileDialog();
            save.Title = "Ubicación del archivo SAP";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;
                erse.CargaDatosSAP(erse.usarFechaFactura, erse.fd, erse.fh);
                erse.ExportExcelSAP(save.FileName);
                Cursor = Cursors.Default;
                MessageBox.Show("Informe terminado.",
                    "Exportación a Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

            }
               

            


        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            //using (var fbd = new FolderBrowserDialog())
            //{
            //    DialogResult result = fbd.ShowDialog();

            //    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            //    {
            //        this.ExportExcel("Prueba.xlsx");
            //    }
            //}

            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del archivo MT";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if(result == DialogResult.OK)
            {
                erse.ExportExcel2(save.FileName, "MT");
            }

            save = new SaveFileDialog();
            save.Title = "Ubicación del archivo BTE";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                erse.ExportExcel2(save.FileName, "BTE");
            }

            save = new SaveFileDialog();
            save.Title = "Ubicación del archivo BTN";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                erse.ExportExcel2(save.FileName, "BTN");
            }

            MessageBox.Show("Informe terminado.",
              "Exportación a Excel",
              MessageBoxButtons.OK,
              MessageBoxIcon.Information);
        }


        

        

        private void parametrizaciónToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void parámetrosERSEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters f = new FrmParameters();
            f.Text = "Parámetros E.R.S.E.";
            f.tabla = "erse_param";
            f.esquemaString = "FAC";
            f.Show();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmInformeERSE_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmInformeERSE" ,"N/A");
        }

    }
}
