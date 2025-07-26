using EndesaBusiness.eer;
using EndesaBusiness.factoring;
using EndesaBusiness.facturacion;
using EndesaBusiness.medida;
using EndesaEntity.factoring;
using MS.WindowsAPICodePack.Internal;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmMes12BT : Form
    {
        string factoring = "";
        EndesaBusiness.utilidades.Param p =
            new EndesaBusiness.utilidades.Param("mes12_param", EndesaBusiness.servidores.MySQLDB.Esquemas.FAC);

        EndesaBusiness.factoring.Mes12BT mes12BT;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        

        public FrmMes12BT()
        {
            usage.Start("Facturación", "FrmMes12BT" ,"N/A");
            mes12BT = new EndesaBusiness.factoring.Mes12BT();
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void importarExcelMes12ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string ficheroExcel = "";
            
            DialogResult result = DialogResult.Yes;

            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Seleccione el archivo Excel de Factoring Mes12";
            d.Filter = "Excel files|*.xls*";
            d.Multiselect = false;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                foreach (string fileName in d.FileNames)
                {
                    ficheroExcel = fileName;                    
                    
                    if (fileName.EndsWith("b"))
                    {
                        mes12BT.es_xlsb = true;
                        ConvertXlsbToXlsx(fileName, fileName.Replace(".xlsb", ".xlsx"));
                        ficheroExcel = fileName.Replace(".xlsb", ".xlsx");
                    }

                    p.UpdateParameter("nombre_archivo", ficheroExcel.Replace(@"\", @"\\"));
                    p = new EndesaBusiness.utilidades.Param("mes12_param", EndesaBusiness.servidores.MySQLDB.Esquemas.FAC);

                    mes12BT.fichero_excel_mes12 = ficheroExcel;
                    mes12BT.CargaExcel(ficheroExcel);
                }

                Cursor.Current = Cursors.Default;

                MessageBox.Show("Importacion de Excel finalizada.",                       
                  "Importación Excel Mes12 BT",
                  MessageBoxButtons.OK, MessageBoxIcon.Information);

                LoadData();

            }
        }

        

        private void FrmMes12BT_Load(object sender, EventArgs e)
        {

            LoadData();
        }
        private void LoadData()
        {
            mes12BT = new Mes12BT();
            

            dgv.AutoGenerateColumns = false;
            dgv.DataSource = mes12BT.datos_excel;
            factoring = mes12BT.UltimaAdjudicacion();

            if( factoring == "")
                factoring = "Sin adjudicación cargada";           
            
            lbl_factoring_adjudicacion.Text = "Factoring adjudicación: " +  factoring;
            lbl_total_agrupadas.Text = string.Format("Total adjudicaciones agrupadas: {0:#,##0}", mes12BT.total_registros_agrupadas());
            lbl_total_individuales.Text = string.Format("Total adjudicaciones individuales: {0:#,##0}", mes12BT.total_registros_individuales());

        }

        

        private void importarPrevisiónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Ficheros Excel|*.xlsx";
            d.Multiselect = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    mes12BT.ImportarPrevision(fileName);

                }
                MessageBox.Show("Importación de Previsiones finalizada",
                 "Importación Finalizada",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Information);

                
            }
        }

        private void btn_lanzar_proceso_Click(object sender, EventArgs e)
        {

            
            
            List<string> lista_previsiones = new List<string>();
            List<string> lista_adjudicaciones = new List<string>();

            txtInfo.Clear();                       
                
            txtInfo.AppendText("Completando CUPS" + Environment.NewLine);
                
            mes12BT.Complementa_CUPS();            
            
            foreach (EndesaEntity.factoring.DatosExcel hoja in mes12BT.datos_excel)
            {

                if(!hoja.hoja.Contains("NO "))
                {
                    txtInfo.AppendText("Cruzando adjudicación: " + factoring
                    + " en hoja Excel: " + hoja.hoja
                    + Environment.NewLine);
                    mes12BT.CruzaConAdjudicaciones(hoja.hoja);                            
                }
                        
            }

           

            mes12BT.GeneraRespuesta();

            if (mes12BT.es_xlsb)
            {
                txtInfo.AppendText("Convirtiendo Excel de xlsx a xlsb" 
                        + Environment.NewLine);

                ConvertXlsxToXlsb(mes12BT.fichero_excel_mes12, mes12BT.fichero_excel_mes12.Replace(".xlsx", ".xlsb"));
            }
                

            MessageBox.Show("Proceso completado.",
                "Mes 12 Factoring BT",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

            


        }

        private void resetExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("¿Desea resetear los valores TIPO y REFERENCIA, y dejar el Excel como recien importado?",
               "Factoring Mes12 BT",
                           MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                mes12BT.ResetExcel();

                MessageBox.Show("Reset completado.",
                  "Mes 12 Factoring BT",
                     MessageBoxButtons.OK, MessageBoxIcon.Information);


                
            }
        }

        private void FrmMes12BT_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmMes12BT" ,"N/A");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void btn_importar_adjudicaciones_Click(object sender, EventArgs e)
        {
            bool hay_error = false;

            try
            {

                
                OpenFileDialog d = new OpenFileDialog();
                d.Title = "Seleccione el archivo Excel de Adjudicaciones Mes13";
                d.Filter = "Excel files|*.xlsx";
                d.Multiselect = false;

                if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    foreach (string fileName in d.FileNames)
                    {
                        mes12BT.ImportarAdjudicacion(fileName);
                    }
                    Cursor.Current = Cursors.Default;                    

                }

                    
                
            }catch (Exception ex)
            {

            }
        }

        private void importarAdjudicaciónToolStripMenuItem_Click_2(object sender, EventArgs e)
        {
            bool hay_error = false;

            try
            {
                
                OpenFileDialog d = new OpenFileDialog();
                d.Title = "Seleccione el archivo Excel de Adjudicaciones Mes13";
                d.Filter = "Excel files|*.xlsx";
                d.Multiselect = false;

                if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    foreach (string fileName in d.FileNames)
                    {
                        mes12BT.ImportarAdjudicacion(fileName);
                    }
                    Cursor.Current = Cursors.Default;

                }

                
            }catch (Exception ex)
            {

            }
        }

        public void ConvertXlsbToXlsx(string xlsbFilePath, string xlsxFilePath)
        {
           

            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            var workbook = excelApp.Workbooks.Open(xlsbFilePath);
            workbook.SaveAs(xlsxFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook);

            workbook.Close();
            excelApp.Quit();

        }

        public void ConvertXlsxToXlsb(string xlsxFilePath, string xlsbFilePath)
        {
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            var workbook = excelApp.Workbooks.Open(xlsxFilePath);
            workbook.SaveAs(xlsbFilePath, Excel.XlFileFormat.xlExcel12);

            workbook.Close();
            excelApp.Quit();
        }

        private void importarAdjudicaciónToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            

            try
            {
                
                OpenFileDialog d = new OpenFileDialog();
                d.Title = "Seleccione el archivo Excel de Adjudicaciones Mes13";
                d.Filter = "Excel files|*.xlsx";
                d.Multiselect = false;

                if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    foreach (string fileName in d.FileNames)
                    {
                        mes12BT.ImportarAdjudicacion(fileName);
                    }
                    Cursor.Current = Cursors.Default;

                }

                
            }
            catch (Exception ex)
            {

            }
        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "mes12_param";
            p.esquemaString = "FAC";
            p.cabecera = "Parámetros Mes12";
            p.Show();
        }

        private void importarAdjudicaciónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Seleccione los dos archivos Excel de Adjudicaciones Mes13";
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if(d.FileNames.Count() != 2)
                {
                    MessageBox.Show("No ha seleccionado los 2 archivos de la adjudicación!!!",
                            "Importar adjudicaciones",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    Cursor.Current = Cursors.WaitCursor;
                    foreach (string fileName in d.FileNames)
                    {
                        mes12BT.ImportarAdjudicacion(fileName);
                    }
                    Cursor.Current = Cursors.Default;

                    LoadData();

                }


                
            }
        }
    }
}

