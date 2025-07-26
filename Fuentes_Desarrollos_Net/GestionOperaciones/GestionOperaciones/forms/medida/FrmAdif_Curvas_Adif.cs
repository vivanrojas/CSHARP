using EndesaEntity;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.medida
{
    public partial class FrmAdif_Curvas_Adif : Form
    {

        EndesaBusiness.medida.ExcelCUPS ex;
        List<EndesaEntity.medida.PuntoSuministro> lc = new List<EndesaEntity.medida.PuntoSuministro>();

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmAdif_Curvas_Adif()
        {

            usage.Start("Medida", "FrmAdif_Curvas_Adif" ,"N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = false;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    ex = new EndesaBusiness.medida.ExcelCUPS(fileName);
                    if (!ex.hayError)
                    {
                        btn_generar_excels.Enabled = true;
                        lc = ex.lista_cups;
                        Carga_DGV(lc);
                    }
                    else
                    {
                        MessageBox.Show(ex.descripcion_error,
                           "Importar Excel",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Information);
                    }

                }
            }
        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btn_generar_excels_Click(object sender, EventArgs e)
        {
            DateTime min_date = new DateTime(4999,12,31);
            DateTime max_date = DateTime.MinValue;

            EndesaBusiness.adif.NuevaMedidaADIF medida_adif = new EndesaBusiness.adif.NuevaMedidaADIF();
            SaveFileDialog save;
            DialogResult result;

            if (cmb_formato_salida.SelectedIndex == 1)
            {
                save = new SaveFileDialog();
                save.Title = "Ubicación del informe";
                save.AddExtension = true;
                save.DefaultExt = "xlsx";
                save.Filter = "Ficheros xslx (*.xlsx)|*.*";
                result = save.ShowDialog();
                if (result == DialogResult.OK)
                {

                    Cursor.Current = Cursors.WaitCursor;

                    if (cmb_formato_salida.SelectedIndex == 1)
                    {

                        foreach (EndesaEntity.medida.PuntoSuministro p in lc)
                        {
                            min_date = p.fd < min_date ? p.fd : min_date;
                            max_date = p.fh > max_date ? p.fh : max_date;
                        }

                        medida_adif =
                            new EndesaBusiness.adif.NuevaMedidaADIF(lc.Select(z => z.cups20).ToList(), min_date, max_date);

                        GeneraExcel(save.FileName, medida_adif.dic_activa, medida_adif.dic_reactiva, lc);

                        DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                        if (result2 == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(save.FileName);
                        }
                        Cursor.Current = Cursors.Default;

                    }
                    
                }
                


                
            }
            else
            {
                save = new SaveFileDialog();
                save.Title = "Ubicación del informe";
                save.AddExtension = true;
                save.DefaultExt = "CSV";
                save.Filter = "Ficheros CSV (*.CSV)|*.*";
                result = save.ShowDialog();
                if (result == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    GeneraCSVKronos(save.FileName, lc);
                    MessageBox.Show("Archivo generado", "CSV Generado",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Cursor.Current = Cursors.Default;
                }



            }



        }


        //private List<EndesaEntity.medida.CurvaVector> Curvas(List<EndesaEntity.medida.PuntoSuministro> lc)
        //{

        //}

        private void Carga_DGV(List<EndesaEntity.medida.PuntoSuministro> lista)
        {
            lblTotalRegistros.Text = string.Format("Total Registros: {0:#,##0}", lista.Count);
            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;
        }

        public void GeneraCSVKronos(string rutaFichero, List<EndesaEntity.medida.PuntoSuministro> lc)
        {
            int f = 0;
            int c = 0;
            bool firstOnlyOne = true;
            DateTime fechaHora = new DateTime();
            DateTime min_date = new DateTime();
            DateTime max_date = new DateTime();
            string linea = "";

            f = 0;
            firstOnlyOne = true;

           

            min_date = new DateTime(4999,12,31);
            max_date = DateTime.MinValue;

            foreach (EndesaEntity.medida.PuntoSuministro p in lc)
            {
                min_date = p.fd < min_date ? p.fd : min_date;
                max_date = p.fh > max_date ? p.fh : max_date;
            }


            EndesaBusiness.adif.NuevaMedidaADIF adif = new EndesaBusiness.adif.NuevaMedidaADIF();
            adif.CargaMedidaHoraria(lc.Select(z => z.cups20).ToList(), min_date, max_date);

            FileInfo file = new FileInfo(rutaFichero);
            if (file.Exists)
                file.Delete();

            StreamWriter swa = new StreamWriter(file.FullName, false);

            foreach (EndesaEntity.medida.PuntoSuministro pc in lc)
            {
                List<CurvaCuartoHoraria> p;
                if (adif.dic_cc.TryGetValue(pc.cups20, out p))
                {                

                    foreach(CurvaCuartoHoraria  pp in p)
                    {
                       
                        linea = pp.fecha.Date.ToString("dd/MM/yyyy")
                            + ";" + Convert.ToInt32(pp.fecha.ToString("HH"))
                            + ";" + pp.AE
                            + ";" + pp.AES
                            + ";" + pp.R1
                            + ";" + pp.R2
                            + ";" + pp.R3
                            + ";" + pp.R4
                            + ";" + pp.cups22;

                        swa.WriteLine(linea);
                    }
                }
            }

            swa.Close();









        }


        public void GeneraExcel(string rutaFichero, 
            Dictionary<string, List<EndesaEntity.medida.CurvaDeCarga>> dic_activa,
            Dictionary<string, List<EndesaEntity.medida.CurvaDeCarga>> dic_reactiva,
            List<EndesaEntity.medida.PuntoSuministro> lc)
        {
            int f = 0;
            int c = 0;
            bool firstOnlyOne = true;
            DateTime fechaHora = new DateTime();

            

            f = 0;
            firstOnlyOne = true;

            FileInfo file = new FileInfo(rutaFichero);

            if (file.Exists)
                file.Delete();                       


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(file);            
            var workSheet = excelPackage.Workbook.Worksheets.Add("Curvas ADIF");


            var headerCells = workSheet.Cells[1, 1, 1, 8];
            var headerFont = headerCells.Style.Font;

            var allCells = workSheet.Cells[1, 1, 1, 8];
            var cellFont = allCells.Style.Font;
            cellFont.Bold = true;

            //list_cc = lista.OrderBy(z => z.cups15).ThenBy(z => z.fecha).ToList();

            foreach (EndesaEntity.medida.PuntoSuministro pc in lc)
            {
                List <EndesaEntity.medida.CurvaDeCarga> pp;
                if (dic_activa.TryGetValue(pc.cups20, out pp))
                {
                    if(pp.Count > 0)
                        fechaHora = pp[0].fecha.AddHours(-1);

                    List<EndesaEntity.medida.CurvaDeCarga> pp_reactiva;
                    if (dic_reactiva.TryGetValue(pc.cups20, out pp_reactiva))
                    {

                    }


                    for (int x = 0; x < pp.Count; x++)
                    {
                        

                        #region Cabecera
                        if (firstOnlyOne)
                        {
                            f++;
                            c = 1;
                            workSheet.Cells[f, c].Value = "CUPS20";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;
                            workSheet.Cells[f, c].Value = "FECHA";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;
                            workSheet.Cells[f, c].Value = "HORA";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;
                            workSheet.Cells[f, c].Value = "Energía Activa Horaria (kWh)";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;
                            workSheet.Cells[f, c].Value = "Energía Reactiva Horaria (kVar)";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                            firstOnlyOne = false;

                        }
                        #endregion

                        for (int p = 0; p < pp[x].numPeriodosHorarios; p++)
                        {
                            f++;
                            fechaHora = fechaHora.AddHours(1);


                            #region 23 Periodos        
                            if (pp[x].numPeriodosHorarios == 23 && p > 1)
                            {
                                if (p == 2)
                                    fechaHora = fechaHora.AddHours(1);

                                c = 1;

                                workSheet.Cells[f, c].Value = pc.cups20; c++;
                                workSheet.Cells[f, c].Value = fechaHora.Date;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                                workSheet.Cells[f, c].Value = fechaHora.ToString("HH:mm"); c++;
                                workSheet.Cells[f, c].Value = pp[x].horaria_activa[p + 1];
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                workSheet.Cells[f, c].Value = pp_reactiva[x].horaria_reactiva[p + 1];
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;



                            }
                            #endregion
                            #region 25 Periodos
                            else if (pp[x].numPeriodosHorarios == 25 && p > 1)
                            {
                                if (p == 3)
                                    fechaHora = fechaHora.AddHours(-1);

                                c = 1;

                                workSheet.Cells[f, c].Value = pc.cups20; c++;
                                workSheet.Cells[f, c].Value = fechaHora.Date;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                                workSheet.Cells[f, c].Value = fechaHora.ToString("HH:mm"); c++;
                                workSheet.Cells[f, c].Value = pp[x].horaria_activa[p];
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                workSheet.Cells[f, c].Value = pp_reactiva[x].horaria_reactiva[p];
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                               



                            }
                            #endregion
                            #region 24 Periodos
                            else
                            {

                                c = 1;


                                workSheet.Cells[f, c].Value = pc.cups20; c++;
                                workSheet.Cells[f, c].Value = fechaHora.Date; 
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                                workSheet.Cells[f, c].Value = fechaHora.ToString("HH:mm"); c++;
                                workSheet.Cells[f, c].Value = pp[x].horaria_activa[p];
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                workSheet.Cells[f, c].Value = pp_reactiva[x].horaria_reactiva[p];
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;


                            }

                            #endregion
                        }
                    }
                }
            }



            
            

            allCells = workSheet.Cells[1, 1, f, 11];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:E1"].AutoFilter = true;
            allCells.AutoFitColumns();
            excelPackage.Save();

            
            

        }

        private void FrmAdif_Curvas_Adif_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Medida", "FrmAdif_Curvas_Adif" ,"N/A");
        }

        private void FrmAdif_Curvas_Adif_Load(object sender, EventArgs e)
        {
            cmb_formato_salida.SelectedIndex = 0;
        }
    }
}
