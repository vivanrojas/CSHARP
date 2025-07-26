using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class InformeExcelCurvasSCE
    {
        public InformeExcelCurvasSCE()
        {

        }

        public void GeneraExcel(string cups20, List<EndesaEntity.medida.CurvaCuartoHorariaInformes> lista)
        {
            int f = 0;
            bool firstOnlyOne = true;
            DateTime fechaHora = new DateTime();

            EndesaBusiness.medida.FuentesMedida fm = new EndesaBusiness.medida.FuentesMedida();

            List<EndesaEntity.medida.CurvaCuartoHorariaInformes> list_cc
                = new List<EndesaEntity.medida.CurvaCuartoHorariaInformes>();

            f = 0;
            firstOnlyOne = true;

            FileInfo file = new FileInfo(@"c:\Temp" + @"\" + cups20
                + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
            ExcelPackage excelPackage = new ExcelPackage(file);
            var workSheet = excelPackage.Workbook.Worksheets.Add("CC_1");
            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;

            if (lista.Count > 0)
            {
                list_cc = lista.OrderBy(z => z.cups15).ThenBy(z => z.fecha).ToList();

                for (int x = 0; x < list_cc.Count; x++)
                {
                    fechaHora = list_cc[x].fecha;

                    #region Cabecera
                    if (firstOnlyOne)
                    {
                        f++;
                        workSheet.Cells[f, 1].Value = "Fecha";
                        workSheet.Cells[f, 2].Value = "Hora";
                        workSheet.Cells[f, 3].Value = "Energía Activa (kWh)";
                        workSheet.Cells[f, 4].Value = "Fuente";
                        workSheet.Cells[f, 5].Value = "Estado";
                        firstOnlyOne = false;

                    }
                    #endregion

                    for (int p = 1; p <= list_cc[x].numPeriodos; p++)
                    {
                        f++;
                        #region 23 Periodos        
                        if (list_cc[x].numPeriodos == 23 && p > 2)
                        {
                            if (p == 3)
                                fechaHora = fechaHora.AddHours(1);

                            workSheet.Cells[f, 1].Value = fechaHora.Date;
                            workSheet.Cells[f, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 2].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 3].Value = list_cc[x].value[((p + 1) * 4) - 3] / 4;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = "#,##0";

                            if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                            {
                                workSheet.Cells[f, 4].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                    Convert.ToInt32(list_cc[x].fa[p + 1]));
                            }

                            workSheet.Cells[f, 5].Value = list_cc[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;
                            workSheet.Cells[f, 1].Value = fechaHora.Date;
                            workSheet.Cells[f, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 2].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 3].Value = list_cc[x].value[((p + 1) * 4) - 2] / 4;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = "#,##0";

                            if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                            {
                                workSheet.Cells[f, 4].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                    Convert.ToInt32(list_cc[x].fa[p + 1]));
                            }

                            workSheet.Cells[f, 5].Value = list_cc[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;

                            workSheet.Cells[f, 1].Value = fechaHora.Date;
                            workSheet.Cells[f, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 2].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 3].Value = list_cc[x].value[((p + 1) * 4) - 1] / 4;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = "#,##0";

                            if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                            {
                                workSheet.Cells[f, 4].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                    Convert.ToInt32(list_cc[x].fa[p + 1]));
                            }

                            workSheet.Cells[f, 5].Value = list_cc[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;

                            workSheet.Cells[f, 1].Value = fechaHora.Date;
                            workSheet.Cells[f, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 2].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 3].Value = list_cc[x].value[((p + 1) * 4)] / 4;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = "#,##0";
                            fechaHora = fechaHora.AddMinutes(15);

                            if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                            {
                                workSheet.Cells[f, 4].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                    Convert.ToInt32(list_cc[x].fa[p + 1]));
                            }

                            workSheet.Cells[f, 5].Value = list_cc[x].estado;
                        }
                        #endregion
                        #region 25 Periodos
                        else if (list_cc[x].numPeriodos == 25 && p > 2)
                        {
                            if (p == 4)
                                fechaHora = fechaHora.AddHours(-1);


                            workSheet.Cells[f, 1].Value = fechaHora.Date;
                            workSheet.Cells[f, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 2].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 3].Value = list_cc[x].value[(p * 4) - 3] / 4;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = "#,##0";

                            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            {

                                workSheet.Cells[f, 4].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                    Convert.ToInt32(list_cc[x].fa[p]));
                            }

                            workSheet.Cells[f, 5].Value = list_cc[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;

                            workSheet.Cells[f, 1].Value = fechaHora.Date;
                            workSheet.Cells[f, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 2].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 3].Value = list_cc[x].value[(p * 4) - 2] / 4;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = "#,##0";

                            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                 (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            {

                                workSheet.Cells[f, 4].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                    Convert.ToInt32(list_cc[x].fa[p]));
                            }

                            workSheet.Cells[f, 5].Value = list_cc[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;

                            workSheet.Cells[f, 1].Value = fechaHora.Date;
                            workSheet.Cells[f, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 2].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 3].Value = list_cc[x].value[(p * 4) - 1] / 4;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = "#,##0";

                            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                 (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            {

                                workSheet.Cells[f, 4].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                    Convert.ToInt32(list_cc[x].fa[p]));
                            }

                            workSheet.Cells[f, 5].Value = list_cc[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;

                            workSheet.Cells[f, 1].Value = fechaHora.Date;
                            workSheet.Cells[f, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 2].Value = fechaHora.ToString("HH:mm:ss");

                            workSheet.Cells[f, 3].Value = list_cc[x].value[(p * 4)] / 4;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = "#,##0";
                            fechaHora = fechaHora.AddMinutes(15);

                            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            {
                                workSheet.Cells[f, 4].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]), Convert.ToInt32(list_cc[x].fa[p]));
                            }

                            workSheet.Cells[f, 5].Value = list_cc[x].estado;

                        }
                        #endregion
                        #region 24 Periodos
                        else
                        {

                            workSheet.Cells[f, 1].Value = fechaHora.Date;
                            workSheet.Cells[f, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 2].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 3].Value = list_cc[x].value[(p * 4) - 3] / 4;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = "#,##0";

                            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            {
                                workSheet.Cells[f, 4].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                    Convert.ToInt32(list_cc[x].fa[p]));
                            }

                            workSheet.Cells[f, 5].Value = list_cc[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;

                            workSheet.Cells[f, 1].Value = fechaHora.Date;
                            workSheet.Cells[f, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 2].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 3].Value = list_cc[x].value[(p * 4) - 2] / 4;
                             

                            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            {
                                workSheet.Cells[f, 4].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                    Convert.ToInt32(list_cc[x].fa[p]));
                            }

                            workSheet.Cells[f, 5].Value = list_cc[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;

                            workSheet.Cells[f, 1].Value = fechaHora.Date;
                            workSheet.Cells[f, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 2].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 3].Value = list_cc[x].value[(p * 4) - 1] / 4;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = "#,##0";


                            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            {
                                workSheet.Cells[f, 4].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                    Convert.ToInt32(list_cc[x].fa[p]));
                            }

                            workSheet.Cells[f, 5].Value = list_cc[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;


                            workSheet.Cells[f, 1].Value = fechaHora.Date;
                            workSheet.Cells[f, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 2].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 3].Value = list_cc[x].value[(p * 4)] / 4;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = "#,##0";
                            fechaHora = fechaHora.AddMinutes(15);

                            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            {
                                workSheet.Cells[f, 4].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                    Convert.ToInt32(list_cc[x].fa[p]));
                            }

                            workSheet.Cells[f, 5].Value = list_cc[x].estado;
                        }

                        #endregion
                    }
                }

                var allCells = workSheet.Cells[1, 1, f, 5];
                allCells.AutoFitColumns();
            }

           
            excelPackage.Save();

        }
    }
}
