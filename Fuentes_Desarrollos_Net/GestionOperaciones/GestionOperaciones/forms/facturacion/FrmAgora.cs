using OfficeOpenXml;
using OfficeOpenXml.Style;
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
    public partial class FrmAgora : Form
    {

        List<string> c;

        EndesaBusiness.facturacion.Agora ag;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        public FrmAgora()
        {
            usage.Start("Facturación", "FrmAgora" ,"N/A");
            InitializeComponent();
            InicializaColumnas();
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {



            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                ExportExcel(save.FileName);
                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }
            
        }

        private void ExportExcel(string rutaFichero)
        {
            int fila = 0;
            int columna = 0;

            FileInfo file = new FileInfo(rutaFichero);

            if (file.Exists)
                file.Delete();

            ExcelPackage excelPackage = new ExcelPackage(file);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Resumen");


            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;

            headerFont.Bold = true;


            fila = 1;
            columna = 0;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "PRIMER MES PENDIENTE"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = ""; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Nº PS"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Suma TAM (€)"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Promedio TAM (€)"; columna++;


            //workSheet.Cells[c[columna] + fila].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //workSheet.Cells[c[columna] + fila].Style.Fill.BackgroundColor.SetColor(Color.Green);



            // workSheet.Cells[c[columna] + fila + ":" + c[columna] + (fila + 1)].Merge = true;

            for (int rows = 0; rows < dgv.Rows.Count; rows++)
            {
                fila++;
                columna = 0;
                for (int col = 0; col < dgv.Rows[rows].Cells.Count; col++)
                {
                    PintaRecuadro(excelPackage, fila, columna);
                    if (dgv.Rows[rows].Cells[col].Value != null)
                        workSheet.Cells[c[columna] + fila].Value = dgv.Rows[rows].Cells[col].Value;

                    if (col > 2)
                        workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = "#,##0";

                    columna++;
                }

            }

            var allCells = workSheet.Cells[1, 1, fila, columna];
            var cellFont = allCells.Style.Font;
            //cellFont.SetFromFont(new Font("Calibri", 8));
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:E1"].AutoFilter = true;
            allCells.AutoFitColumns();
            
            workSheet = excelPackage.Workbook.Worksheets.Add("Detalle");
            fila = 1;
            columna = 0;

            headerCells = workSheet.Cells[1, 1, 1, 12];
            headerFont = headerCells.Style.Font;

            headerFont.Bold = true;

            workSheet.Cells[c[columna] + fila].Value = "NIF"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "CLIENTE"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "CCOUNIPS"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "MES"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "ESTADO"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "TIPO"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "TAM"; columna++;
            // workSheet.Cells[c[columna] + fila].Value = "Nombre Gestor"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "Gestor"; columna++;
            //workSheet.Cells[c[columna] + fila].Value = "Apellido 1"; columna++;
            //workSheet.Cells[c[columna] + fila].Value = "Apellido 2"; columna++;
            //workSheet.Cells[c[columna] + fila].Value = "Desc Responsable Territorial"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "Posición Gestor";

            for (int i = 0; i < ag.ld.Count; i++)
            {
                fila++;
                columna = 0;
                workSheet.Cells[c[columna] + fila].Value = ag.ld[i].nif; columna++;
                workSheet.Cells[c[columna] + fila].Value = ag.ld[i].nombre_cliente; columna++;
                workSheet.Cells[c[columna] + fila].Value = ag.ld[i].cups13; columna++;
                workSheet.Cells[c[columna] + fila].Value = ag.ld[i].ultimo_mes_facturado; columna++;
                workSheet.Cells[c[columna] + fila].Value = ag.ld[i].estado_ltp; columna++;
                workSheet.Cells[c[columna] + fila].Value = ag.ld[i].tipo; columna++;
                workSheet.Cells[c[columna] + fila].Value = ag.ld[i].tam; workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = "#,##0"; columna++;
                workSheet.Cells[c[columna] + fila].Value = ag.ld[i].nombreGestor; columna++;
                //workSheet.Cells[c[columna] + fila].Value = ag.ld[i].apellido1; columna++;
                //workSheet.Cells[c[columna] + fila].Value = ag.ld[i].apellido2; columna++;
                // workSheet.Cells[c[columna] + fila].Value = ag.ld[i].desc_Responsable_Territorial; columna++;
                workSheet.Cells[c[columna] + fila].Value = ag.ld[i].subDireccion; 
            }

            allCells = workSheet.Cells[1, 1, fila, columna + 1];
            cellFont = allCells.Style.Font;
            //cellFont.SetFromFont(new Font("Calibri", 8));
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:I1"].AutoFilter = true;
            allCells.AutoFitColumns();
            excelPackage.Save();

            MessageBox.Show("Informe terminado.",
             "Exportación a Excel",
             MessageBoxButtons.OK,
             MessageBoxIcon.Information);
        }

        private void PintaRecuadro(ExcelPackage excelPackage, int fila, int columna)
        {
            var workSheet = excelPackage.Workbook.Worksheets.First();
            workSheet.Cells[c[columna] + fila].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[c[columna] + fila].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[c[columna] + fila].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[c[columna] + fila].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }



        private void InicializaColumnas()
        {
            
            c = new List<string>();
            for (char i = 'A'; i < 'Z'; i++)            
                c.Add(i.ToString());
        }

        private void FrmAgora_Load(object sender, EventArgs e)
        {
            DateTime fdTrabajo = new DateTime();
            DateTime fhTrabajo = new DateTime();

            DateTime mesAnterior = new DateTime();
            

            int anio;
            int mes;
            int dias_del_mes;


            mesAnterior = DateTime.Now.AddMonths(-1);
            anio = mesAnterior.Year;
            mes = mesAnterior.Month;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);
                      

            fhTrabajo = new DateTime(anio, mes, dias_del_mes);
            fdTrabajo = new DateTime(fhTrabajo.Year, fhTrabajo.Month, 1);

            ag = new EndesaBusiness.facturacion.Agora(fdTrabajo, fhTrabajo);
            dgv.AutoGenerateColumns = false;
            //dgv.DataSource = ag.lr.OrderBy(z => z.primer_mes_pdte);
            dgv.DataSource = ag.lr;
            dgvd.AutoGenerateColumns = false;
            dgvd.DataSource = ag.ld;
            lbl_total_cups.Text = string.Format("Total CUPS: {0:#,##0}", ag.ld.Count());

        }

        private void LoadData()
        {
            //dgv.Rows.Clear();
            //dgv.Refresh();
            dgv.AutoGenerateColumns = false;
            dgv.DataSource = ag.lr.ToList();            
            //dgv.DataSource = ag.lr.OrderBy(z => z.primer_mes_pdte).ToList();
            dgvd.AutoGenerateColumns = false;            
            dgvd.DataSource = ag.ld;
            lbl_total_cups.Text = string.Format("Total CUPS: {0:#,##0}", ag.ld.Count());
        }

        private void btnDel_d_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgvd.SelectedRows)
            {
                for(int i = 0; i < ag.ld.Count(); i++)
                {
                    if (row.Cells[2].Value.ToString() == ag.ld[i].cups13)
                    {   
                        ag.ld.RemoveAt(i);
                    }
                }
                // dgvd.Rows.Remove(row);
                
            }
            ag.CreaResumen();
            LoadData();            
        }

        private void importarExcelDetalleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImportarExcel();
        }

        private void btnImportar_Click(object sender, EventArgs e)
        {
            ImportarExcel();
        }

        private void ImportarExcel()
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Archivo Excel|*.xlsx";
            d.Multiselect = false;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ag.ld.Clear();
                foreach (string fileName in d.FileNames)
                {
                    ag.CargaExcel(fileName);
                    if (!ag.hayError)
                    {
                        ag.CreaResumen();
                        LoadData();
                        cmdExcel.Enabled = true;
                    }
                    else
                    {
                        MessageBox.Show(ag.descripcion_error,
                            "Agora - CargaExcel",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }

                }
            }
        }

        private void cerrarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void acercaDeInformeEstadoPuntosÁgoraToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GestionOperaciones.forms.facturacion.FrmAgora_Ayuda f = new facturacion.FrmAgora_Ayuda();
            f.ShowDialog();
        }

        private void FrmAgora_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmAgora" ,"N/A");
        }
    }
}
