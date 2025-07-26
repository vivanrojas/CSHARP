using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace GestionOperaciones.forms.herramientas
{
    public partial class FrmBusquedaArchivos : Form
    {

        List<EndesaEntity.global.MSAccess> lista;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmBusquedaArchivos()
        {
            usage.Start("Herramientas", "FrmBusquedaArchivos" ,"N/A");
            InitializeComponent();
        }


        private void cerrarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmb_buscar_Click(object sender, EventArgs e)
        {

            if (cmb_extension.SelectedIndex > -1)
            {
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                fbd.ShowNewFolderButton = true;
                fbd.Description = "Seleccione la carpeta de búsqueda";
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    Cursor.Current = Cursors.WaitCursor;
                    BuscaEnDirectorio2(fbd.SelectedPath, cmb_extension.Text);
                    Cursor.Current = Cursors.Default;

                    if (lista.Count > 0)
                    {
                        SaveFileDialog save = new SaveFileDialog();
                        save.Title = "Ubicación del informe Excel";
                        save.AddExtension = true;
                        save.DefaultExt = "xlsx";
                        save.Filter = "Ficheros xslx (*.xlsx)|*.*";
                        DialogResult result2 = save.ShowDialog();
                        if (result2 == DialogResult.OK)
                        {
                            Cursor.Current = Cursors.WaitCursor;
                            this.ExportExcel(save.FileName, cmb_extension.Text.Replace("*.", ""));
                            Cursor.Current = Cursors.Default;

                            MessageBox.Show("Informe terminado.",
                               "Informe Excel",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Information);

                            DialogResult result3 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                            if (result3 == DialogResult.Yes)
                                System.Diagnostics.Process.Start(save.FileName);

                        }
                    }
                    else
                    {
                        MessageBox.Show("No se han encontrado archivos con extensión "
                            + cmb_extension.Text + " en la ruta " + fbd.SelectedPath,
                                      "Búsqueda de Archivos",
                                      MessageBoxButtons.OK,
                                      MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("Por favor, seleccione una extensión de la lista",
               "Búsqueda de Archivos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }

        }

        private void BuscaEnDirectorio(string directorio, string extension)
        {

            FileInfo file;

            try
            {


                foreach (string d in Directory.GetDirectories(directorio))
                {

                    foreach (string f in Directory.GetFiles(d))
                    {
                        file = new FileInfo(f);
                        if (file.Extension == extension)
                        {
                            EndesaEntity.global.MSAccess c = new EndesaEntity.global.MSAccess();
                            c.nombre_archivo = file.Name;
                            c.ruta_completa = file.FullName;
                            c.fechaCreaccion = file.CreationTime;
                            c.fechaModificacion = file.LastWriteTime;
                            c.fechaUltimoAcceso = file.LastAccessTime;
                            c.size = (file.Length / 1024);
                            lista.Add(c);
                        }

                    }
                    BuscaEnDirectorio(d, extension);
                }


            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message,
                  "Búsqueda de Archivos",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);

            }
        }

        private void BuscaEnDirectorio2(string directorio, string extension)
        {

            DirectoryInfo DirInfo = new DirectoryInfo(directorio);
            var files = from f in DirInfo.EnumerateFiles(extension, SearchOption.AllDirectories) select f;
            lista.Clear();

            foreach (var file in files)
            {

                EndesaEntity.global.MSAccess c = new EndesaEntity.global.MSAccess();
                c.nombre_archivo = file.Name;
                c.ruta_completa = file.FullName;
                c.fechaCreaccion = file.CreationTime;
                c.fechaModificacion = file.LastWriteTime;
                c.fechaUltimoAcceso = file.LastAccessTime;
                c.size = (file.Length / 1024);
                lista.Add(c);

            }

        }

        private void BuscaEnDirectorio3(string directorio, string extension)
        {

            string strSql = "";

            
            
            DirectoryInfo DirInfo = new DirectoryInfo(directorio);
            var files = from f in DirInfo.EnumerateFiles(extension, SearchOption.AllDirectories) select f;
            lista.Clear();

            foreach (var file in files)
            {

                EndesaEntity.global.MSAccess c = new EndesaEntity.global.MSAccess();
                c.nombre_archivo = file.Name;
                c.ruta_completa = file.FullName;
                c.fechaCreaccion = file.CreationTime;
                c.fechaModificacion = file.LastWriteTime;
                c.fechaUltimoAcceso = file.LastAccessTime;
                c.size = (file.Length / 1024);

                strSql = "GRANT SELECT ON MSysObjects TO Admin;";
                //ac = new EndesaBusiness.servidores.AccessDB(file.FullName);
                //OleDbCommand cmd = new OleDbCommand(strSql, ac.con);
                //cmd.ExecuteNonQuery();

                strSql = "select count(*) as total_registros from MsysObjects where"
                + " MsysObjects.Type = 1 and"
                + " Left$([Name], 1) <> '~' AND Left$([Name], 4) <> 'Msys'"
                + " ORDER BY MsysObjects.Name;";

                //ac = new EndesaBusiness.servidores.AccessDB(file.FullName);
                //cmd = new OleDbCommand(strSql, ac.con);
                //r = cmd.ExecuteReader();
                //while (r.Read())
                //{
                //    c.tablas_locales = Convert.ToInt32(r["total_registros"]);
                //}

                strSql = "select count(*) as total_registros from MsysObjects where"
                + " MsysObjects.Type = 4 and"
                + " Left$([Name], 1) <> '~' AND Left$([Name], 4) <> 'Msys'"
                + " ORDER BY MsysObjects.Name;";
                //cmd = new OleDbCommand(strSql, ac.con);
                //r = cmd.ExecuteReader();
                //while (r.Read())
                //{
                //    c.tablas_externas = Convert.ToInt32(r["total_registros"]);
                //}

                strSql = "select count(*) as total_registros from MsysObjects where"
                + " (MsysObjects.Type = -32768 or MsysObjects.Type = -32761) and"
                + " Left$([Name], 1) <> '~' AND Left$([Name], 4) <> 'Msys'"
                + " ORDER BY MsysObjects.Name;";
                //cmd = new OleDbCommand(strSql, ac.con);
                //r = cmd.ExecuteReader();
                //while (r.Read())
                //{
                //    c.total_modulos = Convert.ToInt32(r["total_registros"]);
                //}

                //ac.CloseConnection();


                lista.Add(c);

            }
            
        }

        private void FrmBusquedaArchivos_Load(object sender, EventArgs e)
        {
            lista = new List<EndesaEntity.global.MSAccess>();
        }

        private void ExportExcel(string fichero, string extension)
        {
            int f = 1;
            int c = 1;

            FileInfo fileInfo = new FileInfo(fichero);

            if (fileInfo.Exists)
                fileInfo.Delete();


            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add(extension);
            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;

            workSheet.View.FreezePanes(2, 1);

            headerFont.Bold = true;
            workSheet.Cells[f, c].Value = "Fichero"; c++;
            workSheet.Cells[f, c].Value = "Ruta Fichero"; c++;
            workSheet.Cells[f, c].Value = "Tamaño (KB)"; c++;
            workSheet.Cells[f, c].Value = "Fecha de creación"; c++;
            workSheet.Cells[f, c].Value = "Fecha último acceso"; c++;
            workSheet.Cells[f, c].Value = "Fecha de modificación"; c++;
            workSheet.Cells[f, c].Value = "Tablas Locales"; c++;
            workSheet.Cells[f, c].Value = "Tablas ODBC"; c++;
            workSheet.Cells[f, c].Value = "VBA"; c++;


            for (int i = 0; i < lista.Count; i++)
            {
                f++;
                c = 1;
                workSheet.Cells[f, c].Value = lista[i].nombre_archivo; c++;
                workSheet.Cells[f, c].Value = lista[i].ruta_completa; c++;
                workSheet.Cells[f, c].Value = lista[i].size; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = lista[i].fechaCreaccion; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = lista[i].fechaUltimoAcceso; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = lista[i].fechaModificacion; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = lista[i].tablas_locales; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = lista[i].tablas_externas; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = lista[i].total_modulos > 0 ? "S" : "N";  c++;                

            }

            var allCells = workSheet.Cells[1, 1, f, 9];
            allCells.AutoFitColumns();
            workSheet.Cells["A1:I1"].AutoFilter = true;

            headerFont.Bold = true;
            excelPackage.Save();

        }

        private void FrmBusquedaArchivos_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Herramientas", "FrmBusquedaArchivos" ,"N/A");
        }
    }
}
