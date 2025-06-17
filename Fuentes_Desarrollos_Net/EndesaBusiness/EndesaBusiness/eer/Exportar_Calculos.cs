using EndesaBusiness.servidores;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Windows.Forms;

namespace EndesaBusiness.eer
{
    class Exportar_Calculos
    {

        EndesaBusiness.utilidades.Param param;
        public Exportar_Calculos()
        {
            param = new utilidades.Param("eer_param", MySQLDB.Esquemas.CON);
        }


        public void GeneraExcelCalculos(EndesaEntity.punto_suministro.PuntoSuministro ps,
           EndesaBusiness.medida.CurvasEER curva, EndesaBusiness.calendarios.Calendario cal, 
           DateTime fd, string fechaHora_gen_factura, 
           EndesaBusiness.eer.Coef_Excesos_Potencia coef_excesos, 
           EndesaBusiness.eer.Precios_Excesos_Potencia precio_excesos,
           double[] vectorExcesos)
        {

            FileInfo fichero;
            DirectoryInfo dirSalida;
            string nombreFicheroCalculos;

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage;

            int c = 0;
            int f = 0;

            try
            {
                nombreFicheroCalculos = ps.cups20 + "_ " + fd.ToString("yyyyMM") + "_"
                    + fechaHora_gen_factura + "_Calculos_" + ".xlsx";
                dirSalida = new DirectoryInfo(param.GetValue("ruta_salida_facturas", DateTime.Now, DateTime.Now));

                if (!dirSalida.Exists)
                    dirSalida.Create();

                fichero = new FileInfo(dirSalida.FullName + "\\" + nombreFicheroCalculos);

                excelPackage = new ExcelPackage(fichero);

                #region Datos CuartoHorarios

                var workSheet = excelPackage.Workbook.Worksheets.Add("Datos CuartoHorarios");

                var headerCells = workSheet.Cells[1, 1, 1, 8];
                var headerFont = headerCells.Style.Font;

                var allCells = workSheet.Cells[1, 1, 1, 8];
                var cellFont = allCells.Style.Font;



                f++;
                c++;

                workSheet.Cells[f, c].Value = "Fecha";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                workSheet.Cells[f, c].Value = "Hora";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                workSheet.Cells[f, c].Value = "Periodo";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                workSheet.Cells[f, c].Value = "Energía Activa (kWh)";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                workSheet.Cells[f, c].Value = "Energía Reactiva R1 (kVArh)";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                workSheet.Cells[f, c].Value = "Energía Reactiva R4 (kVArh)";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                workSheet.Cells[f, c].Value = "Energía Reactiva (kVArh)";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                workSheet.Cells[f, c].Value = "Potencia (kW)";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                workSheet.Cells[f, c].Value = "Potencia Cont.";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                workSheet.Cells[f, c].Value = "Excesos";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                for (int i = 1; i <= cal.numPeriodosMedidaCuartoHorario; i++)
                {
                    c = 1;
                    f++;
                    workSheet.Cells[f, c].Value = curva.curvaCuartoHorariaDias[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                    workSheet.Cells[f, c].Value = curva.curvaCuartoHorariaDias[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortTimePattern; c++;

                    workSheet.Cells[f, c].Value = cal.vectorPeriodosTarifariosCuartoHorarios[i]; c++;

                    workSheet.Cells[f, c].Value = curva.curvaCuartoHorariaActiva[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = curva.curvaCuartoHorariaReactiva_R1[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = curva.curvaCuartoHorariaReactiva_R4[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = curva.curvaCuartoHorariaReactiva[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = curva.curvaCuartoHorariaPotencias[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = ps.potecias_contratadas[cal.vectorPeriodosTarifariosCuartoHorarios[i]];
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = vectorExcesos[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                }

                workSheet.Cells["A1:J1"].AutoFilter = true;

                allCells = workSheet.Cells[1, 1, f, c];
                allCells.AutoFitColumns();

                #endregion

                #region Datos Horarios

                workSheet = excelPackage.Workbook.Worksheets.Add("Datos Horarios");

                f = 0;
                c = 0;
                f++;
                c++;

                workSheet.Cells[f, c].Value = "Fecha";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                workSheet.Cells[f, c].Value = "Hora";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                workSheet.Cells[f, c].Value = "Periodo";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                workSheet.Cells[f, c].Value = "Energía Activa (kWh)";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                workSheet.Cells[f, c].Value = "Precio Energía Activa";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); c++;

                for (int i = 1; i <= cal.numPeriodosMedidaHorario; i++)
                {
                    c = 1;
                    f++;

                    workSheet.Cells[f, c].Value = curva.curvaHorariaDias[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                    workSheet.Cells[f, c].Value = curva.curvaHorariaDias[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortTimePattern; c++;

                    workSheet.Cells[f, c].Value = cal.vectorPeriodosTarifarios[i]; c++;

                    workSheet.Cells[f, c].Value = curva.curvaHorariaActiva[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = ps.precios_energia.precios_periodo[cal.vectorPeriodosTarifarios[i]];
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.000000"; c++;

                }

                workSheet.Cells["A1:E1"].AutoFilter = true;

                allCells = workSheet.Cells[1, 1, f, c];
                allCells.AutoFitColumns();


                #endregion

                workSheet = excelPackage.Workbook.Worksheets.Add("Otros Datos");
                f = 0;
                c = 0;
                f++;
                c++;

                // *******
                // EXCESOS
                // *******

                double parcial = 0;
                double suma = 0;
                bool firstOnly = true;

                workSheet.Cells[1, 1].Value = "Tipo Punto Medida: ";
                workSheet.Cells[1, 2].Value = ps.tipo_punto_medida;
                workSheet.Cells[1, 3].Value = "Tarifa: ";
                workSheet.Cells[1, 4].Value = ps.tarifa.tarifa;

                f = 2;



                for (int periodoTarifario = 1; periodoTarifario <= ps.tarifa.numPeriodosTarifarios; periodoTarifario++)
                {
                    suma = 0;
                    workSheet.Cells[f, 1].Value = "Excesos P" + periodoTarifario;

                    if (fd < new DateTime(2022, 01, 01))
                    {
                        for (int i = 1; i < cal.numPeriodosMedidaCuartoHorario; i++)
                            if (cal.vectorPeriodosTarifariosCuartoHorarios[i] == periodoTarifario)
                                suma = suma + vectorExcesos[i];

                        workSheet.Cells[f, 2].Value = suma;
                        workSheet.Cells[f, 2].Style.Numberformat.Format = "#,##0.000000";

                        workSheet.Cells[f, 3].Value = "Total P" + periodoTarifario;
                        workSheet.Cells[f, 4].Value =
                            Math.Sqrt(suma) * 
                            coef_excesos.GetValorExcesosPotencia(ps.tarifa.tarifa, periodoTarifario) * 
                            Math.Round(coef_excesos.GetValorExcesosPotencia(ps.tarifa.tarifa, periodoTarifario) / 166.386, 4);
                        workSheet.Cells[f, 4].Style.Numberformat.Format = "#,##0.0000";
                    }
                    else if(fd >= new DateTime(2022, 01, 01))
                    {
                        for (int i = 1; i < cal.numPeriodosMedidaCuartoHorario; i++)
                            if (cal.vectorPeriodosTarifariosCuartoHorarios[i] == periodoTarifario)
                                suma = suma + vectorExcesos[i];

                        workSheet.Cells[f, 2].Value = suma;
                        workSheet.Cells[f, 2].Style.Numberformat.Format = "#,##0.000000";
                        workSheet.Cells[f, 3].Value = "Total P" + periodoTarifario;
                        workSheet.Cells[f, 4].Value =
                            Math.Sqrt(suma) *
                            coef_excesos.GetValorExcesosPotencia(ps.tarifa.tarifa, periodoTarifario) *
                            precio_excesos.GetValorPrecioExcesosPotencia(ps.tarifa.tarifa, ps.tipo_punto_medida);
                        workSheet.Cells[f, 4].Style.Numberformat.Format = "#,##0.0000";

                    }
                   
                    f++;
                }


                for (int i = 1; i <= ps.tarifa.numPeriodosTarifarios; i++)
                {
                    workSheet.Cells[f, 1].Value = "Coef. Excesos (kp) P" + i; 
                    workSheet.Cells[f, 2].Value = coef_excesos.GetValorExcesosPotencia(ps.tarifa.tarifa, i);
                    workSheet.Cells[f, 2].Style.Numberformat.Format = "#,##0.000000"; 
                    f++;

                }

                if (fd < new DateTime(2022, 01, 01))
                {
                    double constanteExcesos = Convert.ToDouble(param.GetValue("constanteExcesos", DateTime.Now, DateTime.Now));
                    workSheet.Cells[3, 6].Value = "Constante Excesos:";
                    workSheet.Cells[3, 7].Value = Math.Round(constanteExcesos / 166.386, 4);
                    workSheet.Cells[3, 7].Style.Numberformat.Format = "#,##0.0000";
                }
                else if (fd >= new DateTime(2022, 01, 01))
                {
                    double precioExcesos = precio_excesos.GetValorPrecioExcesosPotencia(ps.tarifa.tarifa, ps.tipo_punto_medida);
                    workSheet.Cells[3, 6].Value = "Precio Excesos (tep):";
                    workSheet.Cells[3, 7].Value = precioExcesos;
                    workSheet.Cells[3, 7].Style.Numberformat.Format = "#,##0.0000";
                }



                if (fd < new DateTime(2021, 06, 01))
                {
                    for (int pt = 1; pt <= ps.tarifa.numPeriodosTarifarios; pt++)
                    {
                        //parcial = CalculaExceso(cal, pt, ps);
                        if (parcial > 0)
                        {

                            //facturaDetalle.descripcion = facturaDetalle.descripcion + "AC" + pt + ": "
                            //+ string.Format("{0:#,##0.000}", vectorAci[pt]);

                            suma = suma + parcial;
                        }

                    }
                }
                else if (fd > new DateTime(2021, 06, 01) && fd < new DateTime(2022, 01, 01))
                {
                    firstOnly = true;

                    if (curva.curvaCompleta)
                    {
                        for (int pt = 1; pt <= ps.tarifa.numPeriodosTarifarios; pt++)
                        {
                            //parcial = CalculaExceso_anterior_2022_01_01(cal, ps.tarifa.tarifa, pt, ps);
                            if (parcial > 0)
                            {
                                
                                
                                   // facturaDetalle.descripcion = facturaDetalle.descripcion + " AC" + pt + ": "
                                   //+ string.Format("{0:#,##0.000}", vectorAci[pt]);
                                


                                suma = suma + parcial;
                            }
                        }
                    }
                    else
                    {
                        //suma = peaje.importe_excesos_potencia;
                    }



                }
                else // Para calculos a partir del 2022_01_01
                {
                    firstOnly = true;

                    if (curva.curvaCompleta)
                    {
                        for (int pt = 1; pt <= ps.tarifa.numPeriodosTarifarios; pt++)
                        {
                            //parcial = CalculaExceso_posterior_2022_01_01(cal, ps.tarifa.tarifa, pt, ps);
                            if (parcial > 0)
                            {
                                if (firstOnly)
                                {
                                    //facturaDetalle.descripcion = facturaDetalle.descripcion + "AC" + pt + ": "
                                    //+ string.Format("{0:#,##0.000}", vectorAci[pt]);
                                    firstOnly = false;
                                }
                                else
                                {
                                   // facturaDetalle.descripcion = facturaDetalle.descripcion + " AC" + pt + ": "
                                   //+ string.Format("{0:#,##0.000}", vectorAci[pt]);
                                }


                                suma = suma + parcial;
                            }
                        }
                    }
                    else
                    {
                        //suma = peaje.importe_excesos_potencia;
                    }
                }




                allCells = workSheet.Cells[1, 1, 20, 20];
                allCells.AutoFitColumns();






                excelPackage.Save();
            }
            catch (Exception e)
            {
                MessageBox.Show("Medida incompleta para el punto: " + ps.cups20,
                      "Medida Incompleta",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Error);
            }


        }

        

        //private double CalculaExceso_anterior_2022_01_01(EndesaBusiness.calendarios.Calendario cal, string tarifa, int periodoTarifario,
        //    EndesaEntity.punto_suministro.PuntoSuministro ps)
        //{
        //    double total = 0;
        //    double suma = 0;
        //    double[] vectorExcesos;


        //    vectorExcesos = CargaExcesos(cal, ps);

        //    double constanteExcesos = Convert.ToDouble(param.GetValue("constanteExcesos", DateTime.Now, DateTime.Now));


        //    for (int i = 1; i < cal.numPeriodosMedidaCuartoHorario; i++)
        //        if (cal.vectorPeriodosTarifariosCuartoHorarios[i] == periodoTarifario)
        //            suma = suma + vectorExcesos[i];

        //    if (suma > 0)
        //    {
        //        total = Math.Sqrt(suma) * coef_excesos.GetValorExcesosPotencia(tarifa, periodoTarifario) * Math.Round(constanteExcesos / 166.386, 4);
        //        if (ps.tarifa.tarifa.Substring(0, 1) != "2")
        //            vectorAci[periodoTarifario] = Math.Sqrt(suma);
        //    }
        //    else
        //    {
        //        total = 0;
        //    }


        //    return total;
        //}

        //private double CalculaExceso_posterior_2022_01_01(EndesaBusiness.calendarios.Calendario cal,
        //   string tarifa, int periodoTarifario, EndesaEntity.punto_suministro.PuntoSuministro ps)
        //{
        //    double total = 0;
        //    double suma = 0;
        //    double kp = 0;
        //    double tep = 0;


        //    double[] vectorPotenciasMaximasRegistradas;
        //    double[] vectorExcesos;


        //    kp = coef_excesos.GetValorExcesosPotencia(tarifa, periodoTarifario);
        //    tep = precio_excesos.GetValorPrecioExcesosPotencia(tarifa, ps.tipo_punto_medida);

        //    if (ps.tipo_punto_medida < 4)
        //    {
        //        vectorExcesos = CargaExcesos(cal, ps);

        //        for (int i = 1; i < cal.numPeriodosMedidaCuartoHorario; i++)
        //            if (cal.vectorPeriodosTarifariosCuartoHorarios[i] == periodoTarifario)
        //                suma = suma + vectorExcesos[i];

        //        if (suma > 0)
        //        {
        //            total = kp * tep * Math.Sqrt(suma);
        //            vectorAci[periodoTarifario] = Math.Sqrt(suma);
        //        }
        //        else
        //        {
        //            total = 0;
        //        }
        //    }
        //    else
        //    {
        //        vectorPotenciasMaximasRegistradas = CargaPotenciasMaximasRegistradas(cal, ps);

        //    }

        //    return total;
        //}

        //private double CalculaExceso(EndesaBusiness.calendarios.Calendario cal, int periodoTarifario,
        //   EndesaEntity.punto_suministro.PuntoSuministro ps)
        //{
        //    double total = 0;
        //    double suma = 0;


        //    double[] vectorExcesos = CargaExcesos(cal, ps);
        //    double constanteExcesos = Convert.ToDouble(param.GetValue("constanteExcesos", DateTime.Now, DateTime.Now));

        //    for (int i = 1; i < cal.numPeriodosMedidaCuartoHorario; i++)
        //        if (cal.vectorPeriodosTarifariosCuartoHorarios[i] == periodoTarifario)
        //            suma = suma + vectorExcesos[i];

        //    if (suma > 0)
        //    {
        //        total = Math.Sqrt(suma) * coef_excesos.GetValorExcesosPotencia("global", periodoTarifario) * Math.Round(constanteExcesos / 166.386, 4);
        //        if (ps.tarifa.tarifa.Substring(0, 1) != "2" &&
        //            ps.tarifa.tarifa.Substring(0, 1) != "3")
        //            vectorAci[periodoTarifario] = Math.Sqrt(suma);
        //    }
        //    else
        //    {
        //        total = 0;
        //        // vectorAci[periodoTarifario] = 0;
        //    }


        //    return total;
        //}

    }
}
