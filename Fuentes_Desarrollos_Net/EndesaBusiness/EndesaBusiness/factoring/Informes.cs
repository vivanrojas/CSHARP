using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.factoring
{
    public class Informes
    {

        logs.Log ficheroLog;
        private EndesaBusiness.utilidades.Param param;
        public Informes() 
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Mes13_Prevision");
            param = new utilidades.Param("ff_param", servidores.MySQLDB.Esquemas.FAC);
        }


        public void Informe_Estimacion(string rutaFichero, string factoring, DateTime fecha,
            Dictionary<string, EndesaEntity.factoring.Estimacion> dic_inv,
            Dictionary<string, EndesaEntity.factoring.Estimacion> dic_agr,
            Dictionary<string, EndesaEntity.factoring.ListaNegra> dic_lista_negra,
            List<EndesaEntity.factoring.CalendarioFactoring> lista_bloques)
        {

            int contador_individuales = 0;
            string contador_individuales_texto = "";
            int contador_agrupadas = 0;
            string contador_agrupadas_texto = "";

            double numVeces = 0;
            double totalf = 0;

            DateTime fechaFacturaEstimada = new DateTime();
            DateTime fechaFactura = new DateTime();
            utilidades.Fechas utilFechas = new utilidades.Fechas();

            ficheroLog.Add("Generando Informe Excel de Estimación Factorigin " + factoring);
            DateTime fechaFactoring = new DateTime();
            fechaFactoring = new DateTime(Convert.ToInt32(factoring.Substring(0, 4)),
                            Convert.ToInt32(factoring.Substring(4, 2)), 1);



            ficheroLog.Add("Generado Excel --> " + rutaFichero);
            FileInfo file = new FileInfo(rutaFichero);

            int c = 0;
            int f = 0;


            if (file.Exists)
                file.Delete();

            FileInfo plantillaExcel =
                        new FileInfo(System.Environment.CurrentDirectory +
                        param.GetValue("plantilla_prevision"));


            FileInfo fileInfo = new FileInfo(rutaFichero);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);

            var workSheet = excelPackage.Workbook.Worksheets["INDIVIDUALES"];

            var headerCells = workSheet.Cells[1, 1, 1, 30];
            var headerFont = headerCells.Style.Font;

            #region Cebecera_Individuales
            f = 1;
            c = 1;
            workSheet.Cells[f, c].Value = "LN"; c++;
            workSheet.Cells[f, c].Value = "EMPRESA TITULAR"; c++;
            workSheet.Cells[f, c].Value = "NIF"; c++;
            workSheet.Cells[f, c].Value = "CLIENTE"; c++;
            workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
            workSheet.Cells[f, c].Value = "CUPSREE"; c++;
            workSheet.Cells[f, c].Value = "REFERENCIA"; c++;
            workSheet.Cells[f, c].Value = "SEC"; c++;
            workSheet.Cells[f, c].Value = "CONTROL"; c++;
            workSheet.Cells[f, c].Value = "Estimación Importe"; c++;
            workSheet.Cells[f, c].Value = "Estimación Base"; c++;
            workSheet.Cells[f, c].Value = "Estimación Impuestos"; c++;
            workSheet.Cells[f, c].Value = "DiaF"; c++;
            workSheet.Cells[f, c].Value = "DiaV(F+)"; c++;

            

            workSheet.Cells[f, c].Value = "ifactura_"
                + utilFechas.MesTexto(lista_bloques[0].consumos_desde).Substring(0, 3)
                + lista_bloques[0].consumos_desde.ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "ifactura_"
                + utilFechas.MesTexto(lista_bloques[1].consumos_desde).Substring(0, 3)
                + lista_bloques[1].consumos_desde.ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "ifactura_"
                + utilFechas.MesTexto(lista_bloques[2].consumos_desde).Substring(0, 3)
                + lista_bloques[2].consumos_desde.ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "ifactura_"
                + utilFechas.MesTexto(lista_bloques[3].consumos_desde).Substring(0, 3)
                + lista_bloques[3].consumos_desde.ToString("yy"); c++;

            workSheet.Cells[f, c].Value = "impuestos_"
                + utilFechas.MesTexto(lista_bloques[0].consumos_desde).Substring(0, 3)
                + lista_bloques[0].consumos_desde.ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "impuestos_"
                + utilFechas.MesTexto(lista_bloques[1].consumos_desde).Substring(0, 3)
                + lista_bloques[1].consumos_desde.ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "impuestos_"
                + utilFechas.MesTexto(lista_bloques[2].consumos_desde).Substring(0, 3)
                + lista_bloques[2].consumos_desde.ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "impuestos_"
                + utilFechas.MesTexto(lista_bloques[3].consumos_desde).Substring(0, 3)
                + lista_bloques[3].consumos_desde.ToString("yy"); c++;

            workSheet.Cells[f, c].Value = "F_"
                + utilFechas.MesTexto(lista_bloques[1].consumos_desde).Substring(0, 3)
                + lista_bloques[1].consumos_desde.ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "F_"
                + utilFechas.MesTexto(lista_bloques[2].consumos_desde).Substring(0, 3)
                + lista_bloques[2].consumos_desde.ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "F_"
                + utilFechas.MesTexto(lista_bloques[3].consumos_desde).Substring(0, 3)
                + lista_bloques[3].consumos_desde.ToString("yy"); c++;

            workSheet.Cells[f, c].Value = "VTO_"
                + utilFechas.MesTexto(lista_bloques[1].consumos_desde).Substring(0, 3)
                + lista_bloques[1].consumos_desde.ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "VTO_"
                + utilFechas.MesTexto(lista_bloques[2].consumos_desde).Substring(0, 3)
                + lista_bloques[2].consumos_desde.ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "VTO_"
                + utilFechas.MesTexto(lista_bloques[3].consumos_desde).Substring(0, 3)
                + lista_bloques[3].consumos_desde.ToString("yy"); c++;



            headerFont.Bold = true;
            #endregion


            #region Pegado_Excel_Individuales
            foreach (KeyValuePair<string, EndesaEntity.factoring.Estimacion> p in dic_inv)
            {
                f++;
                c = 1;

                contador_individuales = contador_individuales + 1;
                contador_individuales_texto = contador_individuales.ToString().PadLeft(5, '0');

                workSheet.Cells[f, c].Value = p.Value.ln; c++;
                workSheet.Cells[f, c].Value = p.Value.empresa_titular; c++;
                workSheet.Cells[f, c].Value = p.Value.nif; c++;
                workSheet.Cells[f, c].Value = p.Value.cliente; c++;
                if (p.Value.ccounips != null)
                    workSheet.Cells[f, c].Value = p.Value.ccounips; c++;
                if (p.Value.cupsree != null)
                    workSheet.Cells[f, c].Value = p.Value.cupsree; c++;

                workSheet.Cells[f, c].Value = p.Value.factoring.ToString() + "_NR" + contador_individuales_texto; c++;
                workSheet.Cells[f, c].Value = p.Value.sec; c++;
                workSheet.Cells[f, c].Value = p.Value.control; c++;


                #region Estimaciones                
                for (int i = 0; i < 4; i++)
                {
                    if (p.Value.ifactura[i] > 0)
                    {

                        if (i == 0)
                        {
                            if (p.Value.ifactura[i] > 0)
                                numVeces = 2;
                            p.Value.estimacion_importe = p.Value.estimacion_importe + (p.Value.ifactura[i] + p.Value.ifactura[i]);
                        }
                        else
                        {
                            if (p.Value.ifactura[i] > 0)
                                numVeces++;
                            p.Value.estimacion_importe = p.Value.estimacion_importe + p.Value.ifactura[i];
                        }

                    }
                }

                if (numVeces > 0)
                    p.Value.estimacion_importe = Math.Round((p.Value.estimacion_importe / numVeces), 2);

                numVeces = 0;
                totalf = 0;
                for (int i = 0; i < 4; i++)
                {
                    if (p.Value.impuestos[i] > 0)
                    {
                        if (i == 0)
                        {
                            if (p.Value.impuestos[i] > 0)
                            {
                                numVeces = 2;
                                p.Value.estimacion_impuestos = p.Value.estimacion_impuestos + (p.Value.impuestos[i] + p.Value.impuestos[i]);
                            }
                        }
                        else
                        {
                            if (p.Value.impuestos[i] > 0)
                            {
                                numVeces++;
                                p.Value.estimacion_impuestos = p.Value.estimacion_impuestos + p.Value.impuestos[i];
                            }
                        }
                    }
                }


                if (numVeces > 0)
                    p.Value.estimacion_impuestos = Math.Round((p.Value.estimacion_impuestos / numVeces), 2);

                if (p.Value.estimacion_importe > 0 && p.Value.estimacion_impuestos > 0)
                    p.Value.estimacion_base = p.Value.estimacion_importe - p.Value.estimacion_impuestos;

                #endregion


                if (p.Value.estimacion_importe > 0)
                    workSheet.Cells[f, c].Value = p.Value.estimacion_importe; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                if (p.Value.estimacion_base > 0)
                    workSheet.Cells[f, c].Value = p.Value.estimacion_base; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                if (p.Value.estimacion_impuestos > 0)
                    workSheet.Cells[f, c].Value = p.Value.estimacion_impuestos; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;


                // Para dias factura

                numVeces = 0;
                totalf = 0;
                for (int i = 0; i < 4; i++)
                {
                    totalf += p.Value.f[i];
                    if (p.Value.f[i] > 0)
                        numVeces++;
                }
                if (numVeces > 0)
                {
                    totalf = Convert.ToInt32(Math.Round((totalf / numVeces), 0, MidpointRounding.AwayFromZero));
                    fechaFacturaEstimada = fechaFactoring.AddMonths(1);
                    fechaFacturaEstimada = fechaFacturaEstimada.AddDays(totalf - 1);
                    fechaFacturaEstimada = fechaFacturaEstimada.AddDays(RestaDiasExcel1492());
                    fechaFactura = fechaFacturaEstimada;
                    workSheet.Cells[f, c].Value = fechaFacturaEstimada;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }

                c++;

                // Para dias vto
                numVeces = 0;
                totalf = 0;
                for (int i = 0; i < 4; i++)
                {
                    totalf += p.Value.dvto[i];
                    if (p.Value.dvto[i] > 0)
                        numVeces++;
                }
                if (numVeces > 0 && totalf > 0)
                {
                    totalf = Convert.ToInt32(Math.Round((totalf / numVeces), 0, MidpointRounding.AwayFromZero));
                    workSheet.Cells[f, c].Value = fechaFacturaEstimada.AddDays(totalf);
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                else
                {
                    // Si no tenemos fecha le sumamos 2 meses
                    workSheet.Cells[f, c].Value = fechaFactura.AddMonths(2);
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }

                c++;


                numVeces = 0;
                for (int i = 0; i < 4; i++)
                {
                    if (p.Value.ifactura[i] > 0)
                        workSheet.Cells[f, c].Value = p.Value.ifactura[i]; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";

                    c++;
                }

                for (int i = 0; i < 4; i++)
                {
                    if (p.Value.impuestos[i] > 0)
                        workSheet.Cells[f, c].Value = p.Value.impuestos[i]; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    c++;
                }

                for (int i = 1; i < 4; i++)
                {
                    if (p.Value.f[i] > 0)
                        workSheet.Cells[f, c].Value = p.Value.f[i];
                    c++;
                }

                for (int i = 1; i < 4; i++)
                {
                    if (p.Value.dvto[i] > 0)
                        workSheet.Cells[f, c].Value = p.Value.dvto[i];
                    c++;
                }

            }



            var allCells = workSheet.Cells[1, 1, f, 50];
            headerFont = headerCells.Style.Font;

            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:AB1"].AutoFilter = true;
            allCells.AutoFitColumns();


            #endregion
                        
            workSheet = excelPackage.Workbook.Worksheets["AGRUPADAS"];

            headerCells = workSheet.Cells[1, 1, 1, 30];
            headerFont = headerCells.Style.Font;

            #region Cebecera_Agrupadas
            f = 1;
            c = 1;
            workSheet.Cells[f, c].Value = "LN"; c++;
            workSheet.Cells[f, c].Value = "EMPRESA TITULAR"; c++;
            workSheet.Cells[f, c].Value = "NIF"; c++;
            workSheet.Cells[f, c].Value = "CLIENTE"; c++;
            workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
            workSheet.Cells[f, c].Value = "CUPSREE"; c++;
            workSheet.Cells[f, c].Value = "REFERENCIA"; c++;
            workSheet.Cells[f, c].Value = "SEC"; c++;
            workSheet.Cells[f, c].Value = "CONTROL"; c++;
            workSheet.Cells[f, c].Value = "Estimación Importe"; c++;
            workSheet.Cells[f, c].Value = "Estimación Base"; c++;
            workSheet.Cells[f, c].Value = "Estimación Impuestos"; c++;
            workSheet.Cells[f, c].Value = "DiaF"; c++;
            workSheet.Cells[f, c].Value = "DiaV(F+)"; c++;

            workSheet.Cells[f, c].Value = "ifactura_"
                + utilFechas.MesTexto(fechaFactoring.AddYears(-1)).Substring(0, 3)
                + fechaFactoring.AddYears(-1).ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "ifactura_"
                + utilFechas.MesTexto(fechaFactoring.AddMonths(-4)).Substring(0, 3)
                + fechaFactoring.AddMonths(-4).ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "ifactura_"
                + utilFechas.MesTexto(fechaFactoring.AddMonths(-3)).Substring(0, 3)
                + fechaFactoring.AddMonths(-3).ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "ifactura_"
                + utilFechas.MesTexto(fechaFactoring.AddMonths(-2)).Substring(0, 3)
                + fechaFactoring.AddMonths(-2).ToString("yy"); c++;

            workSheet.Cells[f, c].Value = "impuestos_"
                + utilFechas.MesTexto(fechaFactoring.AddYears(-1)).Substring(0, 3)
                + fechaFactoring.AddYears(-1).ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "impuestos_"
                + utilFechas.MesTexto(fechaFactoring.AddMonths(-4)).Substring(0, 3)
                + fechaFactoring.AddMonths(-4).ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "impuestos_"
                + utilFechas.MesTexto(fechaFactoring.AddMonths(-3)).Substring(0, 3)
                + fechaFactoring.AddMonths(-3).ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "impuestos_"
                + utilFechas.MesTexto(fechaFactoring.AddMonths(-2)).Substring(0, 3)
                + fechaFactoring.AddMonths(-2).ToString("yy"); c++;

            workSheet.Cells[f, c].Value = "F_"
                + utilFechas.MesTexto(fechaFactoring.AddMonths(-4)).Substring(0, 3)
                + fechaFactoring.AddMonths(-4).ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "F_"
                + utilFechas.MesTexto(fechaFactoring.AddMonths(-3)).Substring(0, 3)
                + fechaFactoring.AddMonths(-3).ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "F_"
                + utilFechas.MesTexto(fechaFactoring.AddMonths(-2)).Substring(0, 3)
                + fechaFactoring.AddMonths(-2).ToString("yy"); c++;

            workSheet.Cells[f, c].Value = "VTO_"
                + utilFechas.MesTexto(fechaFactoring.AddMonths(-4)).Substring(0, 3)
                + fechaFactoring.AddMonths(-4).ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "VTO_"
                + utilFechas.MesTexto(fechaFactoring.AddMonths(-3)).Substring(0, 3)
                + fechaFactoring.AddMonths(-3).ToString("yy"); c++;
            workSheet.Cells[f, c].Value = "VTO_"
                + utilFechas.MesTexto(fechaFactoring.AddMonths(-2)).Substring(0, 3)
                + fechaFactoring.AddMonths(-2).ToString("yy"); c++;

            headerFont.Bold = true;
            #endregion


            #region Pegado_Excel_Agrupadas

            foreach (KeyValuePair<string, EndesaEntity.factoring.Estimacion> p in dic_agr)
            {
                bool existe_nif = false;
                // Comprobamos que no se repita el NIF
                foreach (KeyValuePair<string, EndesaEntity.factoring.Estimacion> pp in dic_inv)
                {
                    if (pp.Value.nif == p.Value.nif)
                    {
                        existe_nif = true; 
                        break;
                    }
                }
                if (!existe_nif)
                {
                    f++;
                    c = 1;

                    contador_agrupadas = contador_agrupadas + 1;
                    contador_agrupadas_texto = contador_agrupadas.ToString().PadLeft(5, '0');

                    workSheet.Cells[f, c].Value = p.Value.ln; c++;
                    workSheet.Cells[f, c].Value = p.Value.empresa_titular; c++;
                    workSheet.Cells[f, c].Value = p.Value.nif; c++;
                    workSheet.Cells[f, c].Value = p.Value.cliente; c++;
                    if (p.Value.ccounips != null)
                        workSheet.Cells[f, c].Value = p.Value.ccounips; c++;
                    if (p.Value.cupsree != null)
                        workSheet.Cells[f, c].Value = p.Value.cupsree; c++;

                    workSheet.Cells[f, c].Value = p.Value.factoring.ToString() + "_AG" + contador_agrupadas_texto; c++;
                    workSheet.Cells[f, c].Value = p.Value.sec; c++;
                    workSheet.Cells[f, c].Value = p.Value.control; c++;

                    #region Estimaciones
                    numVeces = 0;
                    totalf = 0;
                    //for (int i = 0; i < 4; i++)
                    //{
                    //    if (p.Value.ifactura[i] > 0)
                    //    {
                    //        numVeces++;
                    //        if (i == 0)
                    //            p.Value.estimacion_importe = p.Value.estimacion_importe + (p.Value.ifactura[i] + p.Value.ifactura[i]);
                    //        else
                    //            p.Value.estimacion_importe += p.Value.ifactura[i];
                    //    }
                    //}

                    for (int i = 0; i < 4; i++)
                    {
                        if (p.Value.ifactura[i] > 0)
                        {

                            if (i == 0)
                            {
                                if (p.Value.ifactura[i] > 0)
                                    numVeces = 2;
                                p.Value.estimacion_importe = p.Value.estimacion_importe + (p.Value.ifactura[i] + p.Value.ifactura[i]);
                            }
                            else
                            {
                                if (p.Value.ifactura[i] > 0)
                                    numVeces++;
                                p.Value.estimacion_importe = p.Value.estimacion_importe + p.Value.ifactura[i];
                            }

                        }
                    }

                    if (numVeces > 0)
                        p.Value.estimacion_importe = Math.Round((p.Value.estimacion_importe / numVeces), 2);

                    numVeces = 0;
                    totalf = 0;
                    //for (int i = 0; i < 4; i++)
                    //{
                    //    if (p.Value.impuestos[i] > 0)
                    //    {
                    //        numVeces++;
                    //        if (i == 0)
                    //            p.Value.estimacion_impuestos += (p.Value.impuestos[i] + p.Value.impuestos[i]);
                    //        else
                    //            p.Value.estimacion_impuestos += p.Value.impuestos[i];
                    //    }
                    //}

                    for (int i = 0; i < 4; i++)
                    {
                        if (p.Value.impuestos[i] > 0)
                        {
                            if (i == 0)
                            {
                                if (p.Value.impuestos[i] > 0)
                                {
                                    numVeces = 2;
                                    p.Value.estimacion_impuestos = p.Value.estimacion_impuestos + (p.Value.impuestos[i] + p.Value.impuestos[i]);
                                }
                            }
                            else
                            {
                                if (p.Value.impuestos[i] > 0)
                                {
                                    numVeces++;
                                    p.Value.estimacion_impuestos = p.Value.estimacion_impuestos + p.Value.impuestos[i];
                                }
                            }
                        }
                    }

                    if (numVeces > 0)
                        p.Value.estimacion_impuestos = Math.Round((p.Value.estimacion_impuestos / numVeces), 2);

                    if (p.Value.estimacion_importe > 0 && p.Value.estimacion_impuestos > 0)
                        p.Value.estimacion_base = p.Value.estimacion_importe - p.Value.estimacion_impuestos;


                    #endregion


                    if (p.Value.estimacion_importe > 0)
                        workSheet.Cells[f, c].Value = p.Value.estimacion_importe; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    if (p.Value.estimacion_base > 0)
                        workSheet.Cells[f, c].Value = p.Value.estimacion_base; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    if (p.Value.estimacion_impuestos > 0)
                        workSheet.Cells[f, c].Value = p.Value.estimacion_impuestos; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    numVeces = 0;
                    totalf = 0;
                    for (int i = 0; i < 4; i++)
                    {
                        totalf += p.Value.f[i];
                        if (p.Value.f[i] > 0)
                            numVeces++;
                    }
                    if (numVeces > 0)
                    {
                        totalf = Convert.ToInt32(Math.Round((totalf / numVeces), 0, MidpointRounding.AwayFromZero));
                        fechaFacturaEstimada = fechaFactoring.AddMonths(1);
                        fechaFacturaEstimada = fechaFacturaEstimada.AddDays(totalf - 1);
                        fechaFacturaEstimada = fechaFacturaEstimada.AddDays(RestaDiasExcel1492());
                        fechaFactura = fechaFacturaEstimada;
                        workSheet.Cells[f, c].Value = fechaFacturaEstimada;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }

                    c++;

                    // Para dias vto
                    numVeces = 0;
                    totalf = 0;
                    for (int i = 0; i < 4; i++)
                    {
                        totalf += p.Value.dvto[i];
                        if (p.Value.dvto[i] > 0)
                            numVeces++;
                    }
                    if (numVeces > 0 && totalf > 0)
                    {
                        totalf = Convert.ToInt32(Math.Round((totalf / numVeces), 0, MidpointRounding.AwayFromZero));
                        workSheet.Cells[f, c].Value = fechaFacturaEstimada.AddDays(totalf);
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    else
                    {
                        workSheet.Cells[f, c].Value = fechaFactura;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }

                    c++;

                    for (int i = 0; i < 4; i++)
                    {
                        if (p.Value.ifactura[i] > 0)
                            workSheet.Cells[f, c].Value = p.Value.ifactura[i]; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        c++;
                    }

                    for (int i = 0; i < 4; i++)
                    {
                        if (p.Value.impuestos[i] > 0)
                            workSheet.Cells[f, c].Value = p.Value.impuestos[i]; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        c++;
                    }

                    for (int i = 1; i < 4; i++)
                    {
                        if (p.Value.f[i] > 0)
                            workSheet.Cells[f, c].Value = p.Value.f[i];
                        c++;
                    }

                    for (int i = 1; i < 4; i++)
                    {
                        if (p.Value.dvto[i] > 0)
                            workSheet.Cells[f, c].Value = p.Value.dvto[i];
                        c++;
                    }
                }

            }

            allCells = workSheet.Cells[1, 1, f, 30];
            workSheet.View.FreezePanes(2, 1);
            allCells.AutoFitColumns();
            workSheet.Cells["A1:AB1"].AutoFilter = true;

            #endregion
                        
            workSheet = excelPackage.Workbook.Worksheets["LISTA NEGRA"];

            headerCells = workSheet.Cells[1, 1, 1, 2];
            headerFont = headerCells.Style.Font;

            #region Cabecera_Lista_Negra
            f = 1;
            c = 1;
            workSheet.Cells[f, c].Value = "NIF"; c++;
            workSheet.Cells[f, c].Value = "CLIENTE";
            #endregion

            #region Pegado_Excel_Lista_Negra
            foreach (KeyValuePair<string, EndesaEntity.factoring.ListaNegra> p in dic_lista_negra)
            {
                f++;
                c = 1;
                workSheet.Cells[f, c].Value = p.Key; c++;
                workSheet.Cells[f, c].Value = p.Value.cliente; c++;

            }

            allCells = workSheet.Cells[1, 1, f, 2];
            workSheet.View.FreezePanes(2, 1);
            allCells.AutoFitColumns();
            workSheet.Cells["A1:B1"].AutoFilter = true;
            #endregion            
            excelPackage.SaveAs(fileInfo);

            MessageBox.Show("Informe terminado.",
                               "Previsión Excel",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Information);

            DialogResult result3 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
           MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result3 == DialogResult.Yes)
                System.Diagnostics.Process.Start(rutaFichero);

        }

        private int RestaDiasExcel1492()
        {
            // return -1492
            return 0;
        }
    }
}
