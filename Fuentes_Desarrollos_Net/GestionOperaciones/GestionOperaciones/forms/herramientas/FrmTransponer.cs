using OfficeOpenXml;
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
using EndesaEntity;


namespace GO.forms
{
    public partial class FrmTransponer : Form
    {
        public FrmTransponer()
        {
            InitializeComponent();
        }

        private void btn_transponer_Click(object sender, EventArgs e)
        {
            int totalArchivos = 0;
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Transponer xlsx|*.xlsx";
            d.Multiselect = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                totalArchivos = d.FileNames.Count();
                foreach (string fileName in d.FileNames)
                {
                    Transponer(fileName);
                }
            }
        }

        private void Transponer(string archivo)
        {
            int f = 0;
            int c = 0;
            FileStream fs = new FileStream(archivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            ExcelPackage excelPackage = new ExcelPackage(fs);
            var workSheet = excelPackage.Workbook.Worksheets.First();
            List<EndesaEntity.medida.Medida_Vertical> lista = new List<EndesaEntity.medida.Medida_Vertical>();

            f = 1; // Porque la primera fila es la cabecera
            for (int i = 0; i < 1000000; i++)
            {

                f++;
                if (workSheet.Cells[f, 1].Value == null
                       || workSheet.Cells[f, 2].Value == null
                       || workSheet.Cells[f, 3].Value == null)
                    break;
                else
                {
                    c = 1;
                    EndesaEntity.medida.Medida_Vertical o = new EndesaEntity.medida.Medida_Vertical();
                    o.cups = workSheet.Cells[f, c].Value.ToString().Trim(); c++;
                    o.fecha_numero = Convert.ToInt32(workSheet.Cells[f, c].Value.ToString().Trim()); c++;
                    o.hora = Convert.ToInt32(workSheet.Cells[f, c].Value.ToString().Trim()); c++;
                    o.activa = Convert.ToInt32(workSheet.Cells[f, c].Value.ToString().Trim()); c++;
                    lista.Add(o);
                }
            }

            fs.Close();
            fs = null;
            excelPackage = null;


            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del archivo salida";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                ExportExcel(save.FileName, lista);
                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel generado?", "Abrir Excel transpuesto",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }

        }

        private void ExportExcel(string rutaFichero, List<EndesaEntity.medida.Medida_Vertical> lista)
        {
            int fila = 0;
            int columna = 0;
            bool firstOnly = true;
            Int32 dia = 0;

            FileInfo file = new FileInfo(rutaFichero);

            if (file.Exists)
                file.Delete();

            ExcelPackage excelPackage = new ExcelPackage(file);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Programa Final");

            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;

            List<string> c = new List<string>();
            for (char i = 'A'; i < 'Z'; i++)
                c.Add(i.ToString());


            fila = 1;
            columna = 1;
            
            workSheet.Cells[fila,columna].Value = "cups"; columna++;            
            workSheet.Cells[fila, columna].Value = "Fecha"; columna++;

            columna = 3;

            for(int i = 1; i <= 25; i++)
            {
                workSheet.Cells[fila, columna].Value = "Hora" + i; columna++;
            }                            
                
            
                
            


            
            for(int i = 0; i < lista.Count(); i++)
            {
                
                
                if (firstOnly || (dia != lista[i].fecha_numero))
                {
                    dia = lista[i].fecha_numero;
                    fila++;
                    workSheet.Cells[fila, 1].Value = lista[i].cups; 
                    workSheet.Cells[fila, 2].Value = lista[i].fecha_numero; 
                    firstOnly = false;
                }
                              
               
                columna = lista[i].hora + 3;
                workSheet.Cells[fila, columna].Value = lista[i].activa;
               
            }
           

            var allCells = workSheet.Cells[1, 1, fila, 10];
            allCells.AutoFitColumns();
                      

            headerFont.Bold = true;
            allCells = workSheet.Cells[1, 1, fila, 10];
            allCells.AutoFitColumns();
            excelPackage.Save();

            MessageBox.Show("Informe terminado.",
             "Transponer finalizado",
             MessageBoxButtons.OK,
             MessageBoxIcon.Information);
        }

        private void acercaDeTransponerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.FrmTransponer_Ayuda f = new FrmTransponer_Ayuda();
            //f.ShowDialog();
        }

        private void cerrarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
