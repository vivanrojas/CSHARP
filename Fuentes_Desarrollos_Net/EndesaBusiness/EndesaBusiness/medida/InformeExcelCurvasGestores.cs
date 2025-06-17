using EndesaEntity;
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
    public class InformeExcelCurvasGestores
    {
        public InformeExcelCurvasGestores()
        {

        }

        public void GeneraExcel(string cups20, List<EndesaEntity.medida.CurvaCuartoHorariaInformes> lista)
        {
            int f = 0;
            bool firstOnlyOne = true;
            DateTime fechaHora = new DateTime();

            EndesaBusiness.medida.FuentesMedida fm = new EndesaBusiness.medida.FuentesMedida();

            EndesaBusiness.utilidades.ZipUnZip zip = new EndesaBusiness.utilidades.ZipUnZip();

            List<EndesaEntity.medida.CurvaCuartoHorariaInformes> list_cc 
                = new List<EndesaEntity.medida.CurvaCuartoHorariaInformes>();

            f = 0;
            firstOnlyOne = true;

            FileInfo file = new FileInfo(@"c:\Temp" + @"\" + lista[0].cups22.Substring(0, 20)
                + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
            ExcelPackage excelPackage = new ExcelPackage(file);
            var workSheet = excelPackage.Workbook.Worksheets.Add("CC_1");
            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;

            list_cc = lista.OrderBy(z => z.cups15).ThenBy(z => z.fecha).ToList();

            for (int x = 0; x < list_cc.Count; x++)
            {
                fechaHora = list_cc[x].fecha;

                #region Cabecera
                if (firstOnlyOne)
                {
                    f++;
                    workSheet.Cells[f, 1].Value = "CUPS15";
                    workSheet.Cells[f, 2].Value = "CUPS22";
                    workSheet.Cells[f, 3].Value = "FECHA";
                    workSheet.Cells[f, 4].Value = "HORA";
                    workSheet.Cells[f, 5].Value = "Energía Activa Horaria (kWh)";
                    workSheet.Cells[f, 6].Value = "Energía Reactiva Horaria (kVar)";
                    workSheet.Cells[f, 7].Value = "Potencia Activa";
                    workSheet.Cells[f, 8].Value = "Cuarto Horaria Activa";
                    workSheet.Cells[f, 9].Value = "FUENTE FINAL";
                    workSheet.Cells[f, 10].Value = "ESTADO";

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

                        workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                        workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                        workSheet.Cells[f, 3].Value = fechaHora.Date;
                        workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                        workSheet.Cells[f, 5].Value = list_cc[x].a[p + 1];
                        workSheet.Cells[f, 5].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 6].Value = list_cc[x].r[p + 1];
                        workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 7].Value = list_cc[x].value[((p + 1) * 4) - 3];
                        workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4) - 3] / 4;
                        workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                        if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                            list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                        {
                            workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                Convert.ToInt32(list_cc[x].fa[p + 1]));
                        }

                        workSheet.Cells[f, 10].Value = list_cc[x].estado;

                        fechaHora = fechaHora.AddMinutes(15);
                        f++;
                        workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                        workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                        workSheet.Cells[f, 3].Value = fechaHora.Date;
                        workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                        workSheet.Cells[f, 7].Value = list_cc[x].value[((p + 1) * 4) - 2];
                        workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4) - 2] / 4;
                        workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                        if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                            list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                        {
                            workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                Convert.ToInt32(list_cc[x].fa[p + 1]));
                        }

                        workSheet.Cells[f, 10].Value = list_cc[x].estado;

                        fechaHora = fechaHora.AddMinutes(15);
                        f++;
                        workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                        workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                        workSheet.Cells[f, 3].Value = fechaHora.Date;
                        workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                        workSheet.Cells[f, 7].Value = list_cc[x].value[((p + 1) * 4) - 1];
                        workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4) - 1] / 4;
                        workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                        if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                            list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                        {
                            workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                Convert.ToInt32(list_cc[x].fa[p + 1]));
                        }

                        workSheet.Cells[f, 10].Value = list_cc[x].estado;

                        fechaHora = fechaHora.AddMinutes(15);
                        f++;
                        workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                        workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                        workSheet.Cells[f, 3].Value = fechaHora.Date;
                        workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                        workSheet.Cells[f, 7].Value = list_cc[x].value[((p + 1) * 4)];
                        workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4)] / 4;
                        workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                        fechaHora = fechaHora.AddMinutes(15);

                        if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                            list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                        {
                            workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                Convert.ToInt32(list_cc[x].fa[p + 1]));
                        }

                        workSheet.Cells[f, 10].Value = list_cc[x].estado;
                    }
                    #endregion
                    #region 25 Periodos
                    else if (list_cc[x].numPeriodos == 25 && p > 2)
                    {
                        if (p == 4)
                            fechaHora = fechaHora.AddHours(-1);

                        workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                        workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                        workSheet.Cells[f, 3].Value = fechaHora.Date;
                        workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                        workSheet.Cells[f, 5].Value = list_cc[x].a[p];
                        workSheet.Cells[f, 5].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 6].Value = list_cc[x].r[p];
                        workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 3];
                        workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 3] / 4;
                        workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                        if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                        {

                            workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                Convert.ToInt32(list_cc[x].fa[p]));
                        }

                        workSheet.Cells[f, 10].Value = list_cc[x].estado;

                        fechaHora = fechaHora.AddMinutes(15);
                        f++;
                        workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                        workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                        workSheet.Cells[f, 3].Value = fechaHora.Date;
                        workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                        workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 2];
                        workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 2] / 4;
                        workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                        if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                             (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                        {

                            workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                Convert.ToInt32(list_cc[x].fa[p]));
                        }

                        workSheet.Cells[f, 10].Value = list_cc[x].estado;

                        fechaHora = fechaHora.AddMinutes(15);
                        f++;
                        workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                        workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                        workSheet.Cells[f, 3].Value = fechaHora.Date;
                        workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                        workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 1];
                        workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 1] / 4;
                        workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                        if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                             (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                        {

                            workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                Convert.ToInt32(list_cc[x].fa[p]));
                        }

                        workSheet.Cells[f, 10].Value = list_cc[x].estado;

                        fechaHora = fechaHora.AddMinutes(15);
                        f++;
                        workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                        workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                        workSheet.Cells[f, 3].Value = fechaHora.Date;
                        workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                        workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4)];
                        workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4)] / 4;
                        workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                        fechaHora = fechaHora.AddMinutes(15);

                        if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                        {
                            workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]), Convert.ToInt32(list_cc[x].fa[p]));
                        }

                        workSheet.Cells[f, 10].Value = list_cc[x].estado;

                    }
                    #endregion
                    #region 24 Periodos
                    else
                    {
                        workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                        workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                        workSheet.Cells[f, 3].Value = fechaHora.Date;
                        workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                        workSheet.Cells[f, 5].Value = list_cc[x].a[p];
                        workSheet.Cells[f, 5].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 6].Value = list_cc[x].r[p];
                        workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 3];
                        workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 3] / 4;
                        workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                        if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                        {
                            workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                Convert.ToInt32(list_cc[x].fa[p]));
                        }

                        workSheet.Cells[f, 10].Value = list_cc[x].estado;

                        fechaHora = fechaHora.AddMinutes(15);
                        f++;
                        workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                        workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                        workSheet.Cells[f, 3].Value = fechaHora.Date;
                        workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                        workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 2];
                        workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 2] / 4;
                        workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                        if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                        {
                            workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                Convert.ToInt32(list_cc[x].fa[p]));
                        }

                        workSheet.Cells[f, 10].Value = list_cc[x].estado;

                        fechaHora = fechaHora.AddMinutes(15);
                        f++;
                        workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                        workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                        workSheet.Cells[f, 3].Value = fechaHora.Date;
                        workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                        workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 1];
                        workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 1] / 4;
                        workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";


                        if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                        {
                            workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                Convert.ToInt32(list_cc[x].fa[p]));
                        }

                        workSheet.Cells[f, 10].Value = list_cc[x].estado;

                        fechaHora = fechaHora.AddMinutes(15);
                        f++;
                        workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                        workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                        workSheet.Cells[f, 3].Value = fechaHora.Date;
                        workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                        workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4)];
                        workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4)] / 4;
                        workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                        fechaHora = fechaHora.AddMinutes(15);

                        if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                        {
                            workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                Convert.ToInt32(list_cc[x].fa[p]));
                        }

                        workSheet.Cells[f, 10].Value = list_cc[x].estado;
                    }

                    #endregion
                }
            }

            var allCells = workSheet.Cells[1, 1, f, 11];
            allCells.AutoFitColumns();
            excelPackage.Save();

            zip.ComprimirArchivo(file.FullName, file.FullName.Replace(".xlsx", ".zip"));
            file.Delete();

        }
        public void GeneraExcel(string cups20, Dictionary<string, List<CurvaCuartoHoraria>> dic_cc)
        {
            int f = 0;
            bool firstOnlyOne = true;
            DateTime fechaHora = new DateTime();

            EndesaBusiness.medida.FuentesMedida fm = new EndesaBusiness.medida.FuentesMedida();

            EndesaBusiness.utilidades.ZipUnZip zip = new EndesaBusiness.utilidades.ZipUnZip();

            List<CurvaCuartoHoraria> list_cc = new List<CurvaCuartoHoraria>();

            f = 0;
            firstOnlyOne = true;

            foreach (KeyValuePair<string, List<CurvaCuartoHoraria>> pp in dic_cc)
            {




                FileInfo file = new FileInfo(@"c:\Temp" + @"\" + pp.Value[0].cups22.Substring(0, 20)
                    + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
                ExcelPackage excelPackage = new ExcelPackage(file);
                var workSheet = excelPackage.Workbook.Worksheets.Add("CC_1");
                var headerCells = workSheet.Cells[1, 1, 1, 25];
                var headerFont = headerCells.Style.Font;



                for (int x = 0; x < pp.Value.Count; x++)
                {
                    fechaHora = pp.Value[x].fecha;

                    #region Cabecera
                    if (firstOnlyOne)
                    {
                        f++;
                        workSheet.Cells[f, 1].Value = "CUPS15";
                        workSheet.Cells[f, 2].Value = "CUPS22";
                        workSheet.Cells[f, 3].Value = "FECHA";
                        workSheet.Cells[f, 4].Value = "HORA";
                        workSheet.Cells[f, 5].Value = "Energía Activa Horaria (kWh)";
                        workSheet.Cells[f, 6].Value = "Energía Reactiva Horaria (kVar)";
                        workSheet.Cells[f, 7].Value = "Potencia Activa";
                        workSheet.Cells[f, 8].Value = "Cuarto Horaria Activa";
                        workSheet.Cells[f, 9].Value = "FUENTE FINAL";
                        workSheet.Cells[f, 10].Value = "ESTADO";

                        firstOnlyOne = false;

                    }
                    #endregion

                    for (int p = 1; p <= pp.Value[x].numPeriodos; p++)
                    {
                        f++;
                        #region 23 Periodos        
                        if (pp.Value[x].numPeriodos == 23 && p > 2)
                        {
                            if (p == 3)
                                fechaHora = fechaHora.AddHours(1);

                            workSheet.Cells[f, 1].Value = pp.Value[x].cups15;
                            workSheet.Cells[f, 2].Value = pp.Value[x].cups22;
                            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 5].Value = pp.Value[x].a[p + 1];
                            workSheet.Cells[f, 5].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 6].Value = pp.Value[x].r[p + 1];
                            workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 7].Value = pp.Value[x].value[((p + 1) * 4) - 3];
                            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 8].Value = pp.Value[x].value[((p + 1) * 4) - 3] / 4;
                            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            if (pp.Value[x].fc[p + 1] != null && pp.Value[x].fc[p + 1] != "" &&
                                pp.Value[x].fa[p + 1] != null && pp.Value[x].fa[p + 1] != "")
                            {
                                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(pp.Value[x].fc[p + 1]),
                                    Convert.ToInt32(pp.Value[x].fa[p + 1]));
                            }

                            workSheet.Cells[f, 10].Value = pp.Value[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;
                            workSheet.Cells[f, 1].Value = pp.Value[x].cups15;
                            workSheet.Cells[f, 2].Value = pp.Value[x].cups22;
                            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 7].Value = pp.Value[x].value[((p + 1) * 4) - 2];
                            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 8].Value = pp.Value[x].value[((p + 1) * 4) - 2] / 4;
                            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            if (pp.Value[x].fc[p + 1] != null && pp.Value[x].fc[p + 1] != "" &&
                                pp.Value[x].fa[p + 1] != null && pp.Value[x].fa[p + 1] != "")
                            {
                                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(pp.Value[x].fc[p + 1]),
                                    Convert.ToInt32(pp.Value[x].fa[p + 1]));
                            }

                            workSheet.Cells[f, 10].Value = pp.Value[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;
                            workSheet.Cells[f, 1].Value = pp.Value[x].cups15;
                            workSheet.Cells[f, 2].Value = pp.Value[x].cups22;
                            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 7].Value = pp.Value[x].value[((p + 1) * 4) - 1];
                            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 8].Value = pp.Value[x].value[((p + 1) * 4) - 1] / 4;
                            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            if (pp.Value[x].fc[p + 1] != null && pp.Value[x].fc[p + 1] != "" &&
                                pp.Value[x].fa[p + 1] != null && pp.Value[x].fa[p + 1] != "")
                            {
                                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(pp.Value[x].fc[p + 1]),
                                    Convert.ToInt32(pp.Value[x].fa[p + 1]));
                            }

                            workSheet.Cells[f, 10].Value = pp.Value[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;
                            workSheet.Cells[f, 1].Value = pp.Value[x].cups15;
                            workSheet.Cells[f, 2].Value = pp.Value[x].cups22;
                            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 7].Value = pp.Value[x].value[((p + 1) * 4)];
                            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 8].Value = pp.Value[x].value[((p + 1) * 4)] / 4;
                            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                            fechaHora = fechaHora.AddMinutes(15);

                            if (pp.Value[x].fc[p + 1] != null && pp.Value[x].fc[p + 1] != "" &&
                                pp.Value[x].fa[p + 1] != null && pp.Value[x].fa[p + 1] != "")
                            {
                                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(pp.Value[x].fc[p + 1]),
                                    Convert.ToInt32(pp.Value[x].fa[p + 1]));
                            }

                            workSheet.Cells[f, 10].Value = pp.Value[x].estado;
                        }
                        #endregion
                        #region 25 Periodos
                        else if (pp.Value[x].numPeriodos == 25 && p > 2)
                        {
                            if (p == 4)
                                fechaHora = fechaHora.AddHours(-1);

                            workSheet.Cells[f, 1].Value = pp.Value[x].cups15;
                            workSheet.Cells[f, 2].Value = pp.Value[x].cups22;
                            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 5].Value = pp.Value[x].a[p];
                            workSheet.Cells[f, 5].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 6].Value = pp.Value[x].r[p];
                            workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 7].Value = pp.Value[x].value[(p * 4) - 3];
                            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 8].Value = pp.Value[x].value[(p * 4) - 3] / 4;
                            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            if ((pp.Value[x].fc[p] != null && pp.Value[x].fc[p] != "") &&
                                (pp.Value[x].fa[p] != null && pp.Value[x].fa[p] != ""))
                            {

                                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(pp.Value[x].fc[p]),
                                    Convert.ToInt32(pp.Value[x].fa[p]));
                            }

                            workSheet.Cells[f, 10].Value = pp.Value[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;
                            workSheet.Cells[f, 1].Value = pp.Value[x].cups15;
                            workSheet.Cells[f, 2].Value = pp.Value[x].cups22;
                            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 7].Value = pp.Value[x].value[(p * 4) - 2];
                            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 8].Value = pp.Value[x].value[(p * 4) - 2] / 4;
                            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            if ((pp.Value[x].fc[p] != null && pp.Value[x].fc[p] != "") &&
                                 (pp.Value[x].fa[p] != null && pp.Value[x].fa[p] != ""))
                            {

                                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(pp.Value[x].fc[p]),
                                    Convert.ToInt32(pp.Value[x].fa[p]));
                            }

                            workSheet.Cells[f, 10].Value = pp.Value[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;
                            workSheet.Cells[f, 1].Value = pp.Value[x].cups15;
                            workSheet.Cells[f, 2].Value = pp.Value[x].cups22;
                            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 7].Value = pp.Value[x].value[(p * 4) - 1];
                            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 8].Value = pp.Value[x].value[(p * 4) - 1] / 4;
                            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            if ((pp.Value[x].fc[p] != null && pp.Value[x].fc[p] != "") &&
                                 (pp.Value[x].fa[p] != null && pp.Value[x].fa[p] != ""))
                            {

                                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(pp.Value[x].fc[p]),
                                    Convert.ToInt32(pp.Value[x].fa[p]));
                            }

                            workSheet.Cells[f, 10].Value = pp.Value[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;
                            workSheet.Cells[f, 1].Value = pp.Value[x].cups15;
                            workSheet.Cells[f, 2].Value = pp.Value[x].cups22;
                            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 7].Value = pp.Value[x].value[(p * 4)];
                            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 8].Value = pp.Value[x].value[(p * 4)] / 4;
                            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                            fechaHora = fechaHora.AddMinutes(15);

                            if ((pp.Value[x].fc[p] != null && pp.Value[x].fc[p] != "") &&
                                (pp.Value[x].fa[p] != null && pp.Value[x].fa[p] != ""))
                            {
                                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(pp.Value[x].fc[p]), Convert.ToInt32(pp.Value[x].fa[p]));
                            }

                            workSheet.Cells[f, 10].Value = pp.Value[x].estado;

                        }
                        #endregion
                        #region 24 Periodos
                        else
                        {
                            workSheet.Cells[f, 1].Value = pp.Value[x].cups15;
                            workSheet.Cells[f, 2].Value = pp.Value[x].cups22;
                            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 5].Value = pp.Value[x].a[p];
                            workSheet.Cells[f, 5].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 6].Value = pp.Value[x].r[p];
                            workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 7].Value = pp.Value[x].value[(p * 4) - 3];
                            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 8].Value = pp.Value[x].value[(p * 4) - 3] / 4;
                            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            if ((pp.Value[x].fc[p] != null && pp.Value[x].fc[p] != "") &&
                                (pp.Value[x].fa[p] != null && pp.Value[x].fa[p] != ""))
                            {
                                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(pp.Value[x].fc[p]),
                                    Convert.ToInt32(pp.Value[x].fa[p]));
                            }

                            workSheet.Cells[f, 10].Value = pp.Value[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;
                            workSheet.Cells[f, 1].Value = pp.Value[x].cups15;
                            workSheet.Cells[f, 2].Value = pp.Value[x].cups22;
                            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 7].Value = pp.Value[x].value[(p * 4) - 2];
                            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 8].Value = pp.Value[x].value[(p * 4) - 2] / 4;
                            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            if ((pp.Value[x].fc[p] != null && pp.Value[x].fc[p] != "") &&
                                (pp.Value[x].fa[p] != null && pp.Value[x].fa[p] != ""))
                            {
                                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(pp.Value[x].fc[p]),
                                    Convert.ToInt32(pp.Value[x].fa[p]));
                            }

                            workSheet.Cells[f, 10].Value = pp.Value[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;
                            workSheet.Cells[f, 1].Value = pp.Value[x].cups15;
                            workSheet.Cells[f, 2].Value = pp.Value[x].cups22;
                            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 7].Value = pp.Value[x].value[(p * 4) - 1];
                            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 8].Value = pp.Value[x].value[(p * 4) - 1] / 4;
                            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";


                            if ((pp.Value[x].fc[p] != null && pp.Value[x].fc[p] != "") &&
                                (pp.Value[x].fa[p] != null && pp.Value[x].fa[p] != ""))
                            {
                                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(pp.Value[x].fc[p]),
                                    Convert.ToInt32(pp.Value[x].fa[p]));
                            }

                            workSheet.Cells[f, 10].Value = pp.Value[x].estado;

                            fechaHora = fechaHora.AddMinutes(15);
                            f++;
                            workSheet.Cells[f, 1].Value = pp.Value[x].cups15;
                            workSheet.Cells[f, 2].Value = pp.Value[x].cups22;
                            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            workSheet.Cells[f, 7].Value = pp.Value[x].value[(p * 4)];
                            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, 8].Value = pp.Value[x].value[(p * 4)] / 4;
                            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                            fechaHora = fechaHora.AddMinutes(15);

                            if ((pp.Value[x].fc[p] != null && pp.Value[x].fc[p] != "") &&
                                (pp.Value[x].fa[p] != null && pp.Value[x].fa[p] != ""))
                            {
                                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(pp.Value[x].fc[p]),
                                    Convert.ToInt32(pp.Value[x].fa[p]));
                            }

                            workSheet.Cells[f, 10].Value = pp.Value[x].estado;
                        }

                        #endregion
                    }
                }
                var allCells = workSheet.Cells[1, 1, f, 11];
                allCells.AutoFitColumns();
                excelPackage.Save();

                zip.ComprimirArchivo(file.FullName, file.FullName.Replace(".xlsx", ".zip"));
                file.Delete();
            }

            

        }

    }
}
