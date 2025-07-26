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
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace GestionOperaciones.forms.contratacion.gas
{
    public partial class FrmInformeSolapados : Form
    {
        List<string> c;
        List<EndesaBusiness.contratacion.gestionATRGas.ContratoGasDetalle> lista;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmInformeSolapados()
        {

            usage.Start("Contratación", "FrmInformeSolapados" ,"N/A");

            DateTime fd;
            DateTime fh;
            DateTime mesAnterior;
            int anio;
            int mes;
            int dias_del_mes;            

            mesAnterior = DateTime.Now.AddMonths(-1);
            anio = mesAnterior.Year;
            mes = mesAnterior.Month;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            fh = new DateTime(anio, mes, dias_del_mes);
            fd = new DateTime(fh.Year, fh.Month, 1);

            InitializeComponent();

            txt_fd.Value = fd;
            txt_fh.Value = fh;

            
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            Buscar(txt_fd.Value, txt_fh.Value);
        }

        private void FrmInformeSolapados_Load(object sender, EventArgs e)
        {
            Buscar(txt_fd.Value, txt_fh.Value);
        }

        private void Buscar(DateTime fd, DateTime fh)
        {
            lista = new List<EndesaBusiness.contratacion.gestionATRGas.ContratoGasDetalle>();            
            int numDias = 0;
            int maxDiasContrato = 0;
            DateTime sfd = new DateTime();
            DateTime sfh = new DateTime();

            txtCUPSREE.Text = txtCUPSREE.Text.Trim();

            EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle_Informe inf = 
                new EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle_Informe(fd, fh, (txtCUPSREE.Text.Length > 0) ? txtCUPSREE.Text : null);

            //EndesaBusiness.sigame.Addendas addendas = new EndesaBusiness.sigame.Addendas(fd, fh);
            EndesaBusiness.sigame.Addendas addendas = new EndesaBusiness.sigame.Addendas();
            //addendas.ActualizaAddendas(inf.dic);

            inf.dic = addendas.CompletaInfo(inf.dic, fd);

            // Una vez actualizadas las addendas volvemos a buscar los contratos
            //inf = new EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle_Informe(fd, fh, null);

            foreach (KeyValuePair<string, List<EndesaBusiness.contratacion.gestionATRGas.ContratoGasDetalle>> p in inf.dic)
                if (p.Value.Count > 1)
                {
                    maxDiasContrato = 0;
                    numDias = 0;

                    string cups = p.Value[0].cups20;

                    List<EndesaBusiness.contratacion.gestionATRGas.ContratoGasDetalle> cd =
                        p.Value.FindAll(z => z.fecha_inicio <= fh && (z.fecha_fin >= fd || z.fecha_fin == DateTime.MinValue));

                    cd = cd.OrderBy(z => z.fecha_inicio).ToList();


                    // Buscamos si hay solapamiento
                    for (int i = 0; i < cd.Count; i++)
                    {
                        //if (cd[i].fecha_inicio <= fd && (cd[i].fecha_fin >= fd || cd[i].fecha_fin == DateTime.MinValue))
                        {
                            sfh = cd[i].fecha_fin == DateTime.MinValue ? fh : cd[i].fecha_fin;
                            sfd = cd[i].fecha_inicio < fd ? fd : cd[i].fecha_inicio;
                            sfh = sfh > fh ? fh : sfh;
                            maxDiasContrato = Convert.ToInt32((sfh - sfd).TotalDays + 1) > maxDiasContrato ? Convert.ToInt32((sfh - sfd).TotalDays + 1) : maxDiasContrato;
                            numDias += Convert.ToInt32((sfh - sfd).TotalDays + 1);

                        }
                    }

                    //if(numDias > maxDiasContrato)
                    if (numDias > ((fh - fd).TotalDays + 1))
                        for (int i = 0; i < cd.Count; i++)                    
                            lista.Add(cd[i]);
                    
                }
                    
                        


            lbl_total_contratos.Text = string.Format("Total contratos: {0:#,##0}", lista.Count);
            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;

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
                Cursor.Current = Cursors.WaitCursor;
                //ExportExcel(save.FileName, txt_fd.Value, txt_fh.Value);
                ExportExcel(save.FileName, lista);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }
        }
        private void ExportExcel(string rutaFichero, DateTime fd, DateTime fh)
        {
            bool haySolapamiento = false;
            int numDias = 0;
            DateTime sfd = new DateTime();
            DateTime sfh = new DateTime();

            int fila = 0;
            int columna = 0;

            InicializaColumnas();
            bool firstOnly = true;

            FileInfo file = new FileInfo(rutaFichero);

            if (file.Exists)
                file.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(file);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Solapados");

            var headerCells = workSheet.Cells[1, 1, 1, 10];
            var headerFont = headerCells.Style.Font;

            //fila = 1;
            //workSheet.Cells[c[4] + fila].Value = "PRODUCTO BASE";
            //workSheet.Cells[c[4] + fila + ":" + c[8] + fila].Merge = true;
            //workSheet.Cells[c[4] + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //workSheet.Cells[c[9] + fila].Value = "PRODUCTO SOLAPADO";
            //workSheet.Cells[c[9] + fila + ":" + c[13] + fila].Merge = true;
            //workSheet.Cells[c[9] + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            //headerCells = workSheet.Cells["A2:O2"];
            //headerFont = headerCells.Style.Font;

            fila++;
            columna = 0;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "ESTADO"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "CUPS"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "CLIENTE"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "CIF"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Fecha Inicio"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Fecha Fin"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Qd"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Hora Inicio"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Qi"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Tarifa"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "TIPO"; 
            PintaRecuadro(excelPackage, fila, columna);

            //for (int i = 0; i <= 1; i++)
            //{
            //    PintaRecuadro(excelPackage, fila, columna);
            //    workSheet.Cells[c[columna] + fila].Value = "Fecha Inicio"; columna++;
            //    PintaRecuadro(excelPackage, fila, columna);
            //    workSheet.Cells[c[columna] + fila].Value = "Fecha Fin"; columna++;
            //    PintaRecuadro(excelPackage, fila, columna);
            //    workSheet.Cells[c[columna] + fila].Value = "Qd"; columna++;
            //    PintaRecuadro(excelPackage, fila, columna);
            //    workSheet.Cells[c[columna] + fila].Value = "Tarifa"; columna++;
            //    PintaRecuadro(excelPackage, fila, columna);
            //    workSheet.Cells[c[columna] + fila].Value = "TIPO"; columna++;
            //}

            //PintaRecuadro(excelPackage, fila, columna);
            //workSheet.Cells[c[columna] + fila].Value = "ÚLTIMA MODIFICACIÓN"; columna++;

            EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle_Informe inf = 
                new EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle_Informe(fd, fh, null);

            foreach (KeyValuePair<string, List<EndesaBusiness.contratacion.gestionATRGas.ContratoGasDetalle>> p in inf.dic)
            {
                if (p.Value.Count > 1)
                {
                    numDias = 0;
                    haySolapamiento = false;
                    // Buscamos si hay solapamiento
                    for (int i = 0; i < p.Value.Count; i++)
                    {
                        sfh = p.Value[i].fecha_fin == DateTime.MinValue ? fh : p.Value[i].fecha_fin;
                        sfd = p.Value[i].fecha_inicio < fd ? fd : p.Value[i].fecha_inicio;
                        sfh = sfh > fh ? fh : sfh;
                        numDias += Convert.ToInt32((sfh - sfd).TotalDays + 1);
                    }

                    if (numDias > ((fh - fd).TotalDays + 1))
                    {

                        firstOnly = true;
                        for (int i = 0; i < p.Value.Count; i++)
                        {

                            fila++;
                            columna = 0;
                            if (firstOnly)
                            {
                                workSheet.Cells[c[columna] + fila].Value = p.Value[i].estado; columna++;
                                workSheet.Cells[c[columna] + fila].Value = p.Key; columna++;
                                workSheet.Cells[c[columna] + fila].Value = p.Value[i].customer_name; columna++;
                                workSheet.Cells[c[columna] + fila].Value = p.Value[i].vatnum; columna++;
                                firstOnly = false;
                            }
                            else
                                columna = 4;


                            workSheet.Cells[c[columna] + fila].Value = p.Value[i].fecha_inicio;
                            workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; columna++;

                            if (p.Value[i].fecha_fin > DateTime.MinValue)
                            {
                                workSheet.Cells[c[columna] + fila].Value = p.Value[i].fecha_fin;
                                workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            columna++;


                            workSheet.Cells[c[columna] + fila].Value = p.Value[i].qd;
                            workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = "#,##0"; columna++;

                            if (p.Value[i].hora_inicio > DateTime.MinValue)
                            {
                                workSheet.Cells[c[columna] + fila].Value = p.Value[i].hora_inicio;
                                workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortTimePattern;
                            }
                            columna++;

                            workSheet.Cells[c[columna] + fila].Value = p.Value[i].qi;
                            workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = "#,##0"; columna++;

                            workSheet.Cells[c[columna] + fila].Value = p.Value[i].tarifa; columna++;
                            workSheet.Cells[c[columna] + fila].Value = p.Value[i].tipo.ToUpper(); columna++;
                            //workSheet.Cells[c[columna] + fila].Value = p.Value[i].last_update_date;
                            //workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; columna++;
                        }

                        fila++;
                    }
                }
            }

            var allCells = workSheet.Cells[1, 1, fila, 19];
            allCells.AutoFitColumns();
            workSheet.Cells["A1:I1"].AutoFilter = true;
            workSheet.View.FreezePanes(2, 1);

            headerFont.Bold = true;




            headerFont.Bold = true;
            allCells = workSheet.Cells[1, 1, fila, 9];
            allCells.AutoFitColumns();

            excelPackage.Save();

            MessageBox.Show("Informe terminado.",
             "Exportación a Excel",
             MessageBoxButtons.OK,
             MessageBoxIcon.Information);
        }
        private void ExportExcel(string rutaFichero, List<EndesaBusiness.contratacion.gestionATRGas.ContratoGasDetalle> lista)
        {
            
            int numDias = 0;
            DateTime sfd = new DateTime();
            DateTime sfh = new DateTime();

            int fila = 0;
            int columna = 0;

            InicializaColumnas();
            bool firstOnly = true;

            FileInfo file = new FileInfo(rutaFichero);

            if (file.Exists)
                file.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(file);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Solapados");

            var headerCells = workSheet.Cells[1, 1, 1, 11];
            var headerFont = headerCells.Style.Font;

            

            fila++;
            columna = 0;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "ESTADO"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "CUPS"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "CLIENTE"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "CIF"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Fecha Inicio"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Fecha Fin"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Qd"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Hora Inicio"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Qi"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Tarifa"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "TIPO";
            PintaRecuadro(excelPackage, fila, columna);



            string cups_actual = "";

            foreach (EndesaBusiness.contratacion.gestionATRGas.ContratoGasDetalle p in lista)
            {
                
                        

                fila++;
                columna = 0;

                if(!firstOnly && p.cups20 != cups_actual)
                    fila++;

                if (firstOnly || p.cups20 != cups_actual)
                {
                    cups_actual = p.cups20;
                    workSheet.Cells[c[columna] + fila].Value = p.estado; columna++;
                    workSheet.Cells[c[columna] + fila].Value = p.cups20; columna++;
                    workSheet.Cells[c[columna] + fila].Value = p.customer_name; columna++;
                    workSheet.Cells[c[columna] + fila].Value = p.vatnum; columna++;
                    firstOnly = false;
                }
                else
                    columna = 4;


                workSheet.Cells[c[columna] + fila].Value = p.fecha_inicio;
                workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; columna++;

                if (p.fecha_fin > DateTime.MinValue)
                {
                    workSheet.Cells[c[columna] + fila].Value = p.fecha_fin;
                    workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                columna++;


                workSheet.Cells[c[columna] + fila].Value = p.qd;
                workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = "#,##0"; columna++;

                if (p.hora_inicio > DateTime.MinValue)
                {
                    workSheet.Cells[c[columna] + fila].Value = p.hora_inicio;
                    workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortTimePattern;
                }
                columna++;

                workSheet.Cells[c[columna] + fila].Value = p.qi;
                workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = "#,##0"; columna++;

                workSheet.Cells[c[columna] + fila].Value = p.tarifa; columna++;
                workSheet.Cells[c[columna] + fila].Value = p.tipo.ToUpper(); columna++;
                
                        

            //fila++;
                    
                
            }

            var allCells = workSheet.Cells[1, 1, fila, 19];
            allCells.AutoFitColumns();
            workSheet.Cells["A1:I1"].AutoFilter = true;
            workSheet.View.FreezePanes(2, 1);

            headerFont.Bold = true;




            headerFont.Bold = true;
            allCells = workSheet.Cells[1, 1, fila, 9];
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
            workSheet.Cells[c[columna] + fila].Style.Border.Top.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[c[columna] + fila].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[c[columna] + fila].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[c[columna] + fila].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
        }

        private void InicializaColumnas()
        {

            c = new List<string>();
            for (char i = 'A'; i < 'Z'; i++)
                c.Add(i.ToString());
        }

        private void FrmInformeSolapados_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Contratación", "FrmInformeSolapados" ,"N/A");
        }
    }

}
