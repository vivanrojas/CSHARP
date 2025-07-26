using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class InformeExcelCurvasADIF
    {
        public InformeExcelCurvasADIF()
        {

        }

        public void GeneraExcel(string cups20, List<EndesaEntity.medida.CurvaCuartoHorariaInformes> lista)
        {
            int f = 1;
            int c = 1;
            bool firstOnlyOne = true;
            DateTime fechaHora = new DateTime();

            EndesaBusiness.medida.FuentesMedida fm = new EndesaBusiness.medida.FuentesMedida();

            List<EndesaEntity.medida.CurvaCuartoHorariaInformes> list_cc
                = new List<EndesaEntity.medida.CurvaCuartoHorariaInformes>();

            f = 0;
            firstOnlyOne = true;

            FileInfo file = new FileInfo(@"c:\Temp" + @"\" + cups20
                + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(file);
            var workSheet = excelPackage.Workbook.Worksheets.Add("CH");

            var headerCells = workSheet.Cells[1, 1, 1, 26];
            var headerFont = headerCells.Style.Font;

            
            headerFont.Bold = true;
            
            workSheet.Cells[f, c].Value = "FECHA"; c++;
            workSheet.Cells[f, c].Value = "HORA"; c++;
            workSheet.Cells[f, c].Value = "AE"; c++;
            workSheet.Cells[f, c].Value = "AS"; c++;
            workSheet.Cells[f, c].Value = "R1"; c++;
            workSheet.Cells[f, c].Value = "R2"; c++;
            workSheet.Cells[f, c].Value = "R3"; c++;
            workSheet.Cells[f, c].Value = "R4"; c++;
            workSheet.Cells[f, c].Value = "CUPS22"; c++;


            if (lista.Count > 0)
            {
                list_cc = lista.OrderBy(z => z.cups15).ThenBy(z => z.fecha).ToList();

                for (int x = 0; x < list_cc.Count; x++)
                {
                    fechaHora = list_cc[x].fecha;

                    for (int p = 1; p <= list_cc[x].numPeriodos; p++)
                    {
                        f++;

                    }
                }
                var allCells = workSheet.Cells[1, 1, f, 5];
                allCells.AutoFitColumns();

            }
            excelPackage.Save();
        }


    }
}
