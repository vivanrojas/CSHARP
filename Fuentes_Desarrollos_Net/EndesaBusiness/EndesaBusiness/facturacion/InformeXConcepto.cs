using EndesaBusiness.office;
using EndesaBusiness.utilidades;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace EndesaBusiness.facturacion
{
    public class InformeXConcepto
    {
        public DateTime fd { get; set; }
        public DateTime fh { get; set; }
        public String tipoNegocio { get; set; }
        public String cupsREE { get; set; }
        public String cnifdnic { get; set; }
        public String tipoFactura { get; set; }
        public String conceptos { get; set; }
        public Boolean fechaFactura { get; set; }
        public string empresas { get; set; }
        public string mainQuery { get; set; }

        public List<EndesaEntity.facturacion.FacturaInffact> lf { get; set; }

        public void CargaDatos()
        {
            lf = new List<EndesaEntity.facturacion.FacturaInffact>();
        }

        public void PintarExcel()
        {
            int f = 1;
            int c = 1;

            String[] lista_Conceptos;

            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del archivo Estado medida ADIF";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {

                #region Excel

                lista_Conceptos = conceptos.Split(';');

                FileInfo fileInfo = new FileInfo(save.FileName);

                ExcelPackage excelPackage;
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("Conceptos de facturación");

                var headerCells = workSheet.Cells[1, 1, 1, 15 + lista_Conceptos.Count()];
                var headerFont = headerCells.Style.Font;

                headerFont.Bold = true;

                facturacion.ConceptosFacturacion cf = new ConceptosFacturacion();

                workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                workSheet.Cells[f, c].Value = "CNIFDNIC"; c++;
                workSheet.Cells[f, c].Value = "DAPERSOC"; c++;
                workSheet.Cells[f, c].Value = "CFACTURA"; c++;
                workSheet.Cells[f, c].Value = "FFACTURA"; c++;
                workSheet.Cells[f, c].Value = "FFACTDES"; c++;
                workSheet.Cells[f, c].Value = "FACTHAS"; c++;
                workSheet.Cells[f, c].Value = "VCUOVAFA"; c++;
                workSheet.Cells[f, c].Value = "IVA"; c++;
                workSheet.Cells[f, c].Value = "IFACTURA"; c++;
                workSheet.Cells[f, c].Value = "CREFEREN"; c++;
                workSheet.Cells[f, c].Value = "SECFACTU"; c++;
                workSheet.Cells[f, c].Value = "TESTFACT"; c++;
                workSheet.Cells[f, c].Value = "CUPSREE"; c++;
                workSheet.Cells[f, c].Value = "TFACTURA"; c++;


                

                for (int i = 0; i < lista_Conceptos.Count(); i++)
                {
                    workSheet.Cells[f, c].Value = cf.Descripcion_Corta(Convert.ToInt32(lista_Conceptos[i]));
                    c++;
                }

                for (int i = 0; i < lf.Count(); i++)
                {
                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = lf[i].cemptitu; c++;
                    workSheet.Cells[f, c].Value = lf[i].cnifdnic; c++;
                    workSheet.Cells[f, c].Value = lf[i].dapersoc; c++;
                    workSheet.Cells[f, c].Value = lf[i].cfactura; c++;


                    if (lf[i].ffactura > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = lf[i].ffactura;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if(lf[i].ffactdes > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = lf[i].ffactdes;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;


                    if (lf[i].ffacthas > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = lf[i].ffacthas;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    workSheet.Cells[f, c].Value = lf[i].vcuovafa; c++;
                    workSheet.Cells[f, c].Value = lf[i].iva; c++;
                    workSheet.Cells[f, c].Value = lf[i].ifactura;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = lf[i].creferen; c++;
                    workSheet.Cells[f, c].Value = lf[i].secfactu; c++;
                    workSheet.Cells[f, c].Value = lf[i].testfact; c++;
                    workSheet.Cells[f, c].Value = lf[i].cupsree; c++;
                    workSheet.Cells[f, c].Value = lf[i].tfactura; c++;

                    for (int j = 0; j < lista_Conceptos.Count(); j++)
                    {
                        for (int k = 0; k < 20; k++)
                        {
                            if (lf[i].tconfact[k] == Convert.ToInt32(lista_Conceptos[j]))
                            {
                                workSheet.Cells[f, c + j].Value = lf[i].iconfact[k];
                                workSheet.Cells[f, c + j].Style.Numberformat.Format = "#,##0.00";
                            }
                        }
                    }
                }



                var allCells = workSheet.Cells[1, 1, f, c];
                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells[1, 15 + lista_Conceptos.Count()].AutoFilter = true;
                allCells.AutoFitColumns();
                excelPackage.Save();


                #endregion



                MessageBox.Show("Informe terminado.",
                  "Exportación a Excel",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Information);





                System.Diagnostics.Process.Start(save.FileName);




            }
            

            

          

            
            

           
        }
    }
}
