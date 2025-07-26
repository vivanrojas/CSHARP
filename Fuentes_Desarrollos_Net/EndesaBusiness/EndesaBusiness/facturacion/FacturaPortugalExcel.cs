using EndesaBusiness.servidores;
using OfficeOpenXml;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Globalization;

namespace EndesaBusiness.facturacion
{
    public class FacturaPortugalExcel
    {
        EndesaBusiness.calendarios.UtilidadesCalendario utilFecha;
        utilidades.Param p;
        public FacturaPortugalExcel()
        {
            utilFecha = new EndesaBusiness.calendarios.UtilidadesCalendario();
            p = new utilidades.Param("ag_pt_param", MySQLDB.Esquemas.FAC);
        }


        public void GeneraExcel(string carpeta_cliente, string ruta_archivo, DateTime fd, DateTime fh,
            Spot spot, List<double> cc, List<EndesaEntity.facturacion.ClicksPT> lista_cliks)
        {

            DirectoryInfo dirSalida;
            bool imprimirCalculos = true;

            int c = 0;
            int f = 0;
            int filaCalendario = 0;
            string[] listaHojas = { "CIERRES", "DATOS", "Calculos", "Datos salida", "Factura" };
            List<string> listaPlantillas = new List<string>();

            FileInfo fichero = new FileInfo(ruta_archivo);
            dirSalida = new DirectoryInfo(p.GetValue("ruta_salida_facturas", DateTime.Now, DateTime.Now).Replace("yyyy", fd.Year.ToString())
                + "\\" + fd.Month + "-" + utilFecha.MesLetra(fd) + " " + fd.Year + "\\" + carpeta_cliente);
            if (!dirSalida.Exists)
                dirSalida.Create();

            FileInfo filesave = new FileInfo(dirSalida.FullName + "\\" + fichero.Name);

            //FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);    
            ExcelPackage excelPackage = new ExcelPackage(fichero);
            //var workSheet = excelPackage.Workbook.Worksheets.First();

            #region Cierres
            var workSheet = excelPackage.Workbook.Worksheets[listaHojas[0]];
            f = 1;
            if (lista_cliks != null && false)
                for (int i = 2; i < lista_cliks.Count(); i++)
                {
                    c = 1;
                    f++;
                    workSheet.Cells[f, c].Value = lista_cliks[i].click; c++;
                    workSheet.Cells[f, c].Value = lista_cliks[i].fecha_operacion;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = lista_cliks[i].volumen; c++;
                    workSheet.Cells[f, c].Value = lista_cliks[i].fecha_desde;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = lista_cliks[i].fecha_hasta;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = lista_cliks[i].bl;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    c++; // PL no se informa
                    workSheet.Cells[f, c].Value = lista_cliks[i].fee;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                }




            #endregion

            #region Datos
            workSheet = excelPackage.Workbook.Worksheets[listaHojas[1]];
            #endregion

            #region Calculos
            if (imprimirCalculos)
            {
                workSheet = excelPackage.Workbook.Worksheets[listaHojas[2]];
                // Buscamos la fila desde donde debemos a comenzar a pegar
                for (int fila = 2; fila < 40000; fila++)
                {
                    if (Convert.ToDateTime(workSheet.Cells[fila, 1].Value.ToString()) == fd)
                    {
                        filaCalendario = fila;
                        break;
                    }
                }
                filaCalendario--;
                for (int i = 0; i < spot.lista.Count(); i++)
                {
                    filaCalendario++;
                    c = 6; // Columna
                    workSheet.Cells[filaCalendario, c].Value = spot.lista[i].precio;
                    workSheet.Cells[filaCalendario, c].Style.Numberformat.Format = "#,##0.00"; c++; // Precios Pool
                    workSheet.Cells[filaCalendario, c].Value = cc[i];
                    workSheet.Cells[filaCalendario, c].Style.Numberformat.Format = "#,##0"; c++; // Curva Kwh
                    //workSheet.Cells[filaCalendario, c].Value = ""; c++; // Curva Cerrada
                    //workSheet.Cells[filaCalendario, c].Value = ""; c++; // Curva Abierta
                    //workSheet.Cells[filaCalendario, c].Value = ""; c++; // Precio OE Cerrado
                    //workSheet.Cells[filaCalendario, c].Value = ""; c++; // Precio OE Abierto
                    //workSheet.Cells[filaCalendario, c].Value = ""; c++; // Coste OE Cerrado
                    //workSheet.Cells[filaCalendario, c].Value = ""; c++; // Coste OE Abierto
                }
            }

            #endregion
            workSheet = excelPackage.Workbook.Worksheets[listaHojas[3]];

            #region Datos Salida

            // MES
            workSheet.Cells[3, 13].Value = fd.Month;

            #endregion

            #region Factura
            workSheet = excelPackage.Workbook.Worksheets[listaHojas[4]];

            // Direccion PS


            // Periodo Facturacion
            workSheet.Cells["E28"].Value = "De " + fd.ToString("dd/MM/yyyy") + " a " + fh.ToString("dd/MM/yyyy");

            #endregion


            excelPackage.Save();
            excelPackage.SaveAs(filesave);
            excelPackage = null;

        }

        public void GeneraPDF(string ruta_archivo, DateTime fechaEmision, string numFactura)
        {
            FileInfo fichero = new FileInfo(ruta_archivo);
            string[] listaHojas = { "CIERRES", "DATOS", "Calculos", "Datos salida", "Factura" };

            ExcelPackage excelPackage = new ExcelPackage(fichero);


            var workSheet = excelPackage.Workbook.Worksheets[listaHojas[4]];

            // Fecha Emision
            workSheet.Cells["E27"].Value = fechaEmision.ToString("dd/MM/yyyy");
            workSheet.Cells["E27"].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

            // Num Factura
            workSheet.Cells["E29"].Value = numFactura;
            excelPackage.Save();
            excelPackage = null;

            // Salvado a PDF
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook libro = excel.Workbooks.Open(fichero.FullName);
            Worksheet hoja = libro.Worksheets[listaHojas[4]];
            hoja.Select(Type.Missing);
            // libro.Save();

            libro.SaveAs(fichero.FullName, Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, Type.Missing, Type.Missing,
            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            libro.Close();
            Marshal.ReleaseComObject(libro);
            Marshal.ReleaseComObject(excel);
        }

        public void GeneraExcelDIA(EndesaBusiness.facturacion.InventarioFacturacionPortugal inv,
            EndesaBusiness.facturacion.PagosPortugal pp,
            EndesaBusiness.medida.CCRD medida_cc, DateTime fd, DateTime fh)
        {
            DirectoryInfo dirSalida;
            
            Dictionary<string, int> dic_columna_cup = new Dictionary<string, int>();
            int c = 0;
            int f = 0;
            int filaCalendario = 0;
            
            string cups_columna = "";
            string[] listaHojas = { "CIERRES", "CCH", "COEFICIENTES", "Datos salida", "Distribución_consumo", "Precio TE", "FACTURAR" };
            List<string> listaPlantillas = new List<string>();

            string cups13 = "";
            List<double> cc = new List<double>();
            int o = 0;

            FileInfo fichero = new FileInfo(@"\\e20aemsioa00.enelint.global\M\Facturacion\TOP\TOP1\PLANTILLAS REALES\2020\PORTUGAL\DIA\GRUPO_DIA_2020.xlsx");
            dirSalida = new DirectoryInfo(p.GetValue("ruta_salida_facturas", DateTime.Now, DateTime.Now).Replace("yyyy", fd.Year.ToString())
                + "\\" + fd.Month + "-" + utilFecha.MesLetra(fd) + " " + fd.Year + "\\" + "DIA");
            if (!dirSalida.Exists)
                dirSalida.Create();

            FileInfo filesave = new FileInfo(dirSalida.FullName + "\\" 
                + fichero.Name.Replace(".xlsx","") + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");            
            ExcelPackage excelPackage = new ExcelPackage(fichero);
            

            #region CCH
            var workSheet = excelPackage.Workbook.Worksheets[listaHojas[1]];
            // Buscamos en que columna está cada CUPS
            for (int x = 6; x < 1000; x++)
            {
                if (workSheet.Cells[1, x].Value == null)
                    break;
                else
                {
                    cups_columna = workSheet.Cells[1, x].Value.ToString();
                    if (!dic_columna_cup.TryGetValue(cups_columna, out o))
                        dic_columna_cup.Add(cups_columna, x);
                }
            }

            for (int fila = 2; fila < 40000; fila++)
            {
                if (Convert.ToDateTime(workSheet.Cells[fila, 1].Value.ToString()) == fd)
                {
                    filaCalendario = fila;
                    filaCalendario--;
                    break;
                }
            }

            foreach (KeyValuePair<string, int> p in dic_columna_cup)
            {
                c = p.Value;
                f = filaCalendario;
                cc.Clear();
                cups13 = inv.GetCUPS13FromCUPS20(p.Key);

                if (medida_cc.CurvaCompleta(cups13))
                {
                    cc = medida_cc.GetCurvaVertical(cups13);
                    for (int i = 0; i < cc.Count(); i++)
                    {
                        f++;
                        workSheet.Cells[f, c].Value = cc[i];
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; // Curva Kwh
                    }

                }



            }
            #endregion

            workSheet = excelPackage.Workbook.Worksheets[listaHojas[3]];
            #region Datos Salida

            // MES
            workSheet.Cells[3, 2].Value = fd.Month;
            #endregion



            workSheet = excelPackage.Workbook.Worksheets[listaHojas[6]];
            dic_columna_cup.Clear();
            // Buscamos en que columna está cada CUPS
            for (int x = 2; x < 1000; x++)
            {
                if (workSheet.Cells[1, x].Value == null)
                    break;
                else
                {
                    cups_columna = workSheet.Cells[1, x].Value.ToString();
                    if (!dic_columna_cup.TryGetValue(cups_columna, out o))
                        dic_columna_cup.Add(cups_columna, x);
                }
            }

            foreach (KeyValuePair<string, int> p in dic_columna_cup)
            {
                c = p.Value;                
                List<EndesaEntity.facturacion.ValorConceptoPortugalRAM> oo;
                if(pp.dic.TryGetValue(p.Key, out oo))
                {
                    for (int x = 0; x < oo.Count; x++)
                    {
                        #region P05_P06
                        switch (oo[x].codigo)
                        {
                            case "15150":
                                f = 7;
                                workSheet.Cells[f, c].Value = oo[x].cantidad;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                                workSheet.Cells[f, c].Value = oo[x].precio;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.0000"; f++;
                                workSheet.Cells[f, c].Value = oo[x].valor;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                break;
                            case "15160":
                                f = 10;
                                workSheet.Cells[f, c].Value = oo[x].cantidad;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                                workSheet.Cells[f, c].Value = oo[x].precio;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.0000"; f++;
                                workSheet.Cells[f, c].Value = oo[x].valor;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                break;
                            case "15170":
                                f = 13;
                                workSheet.Cells[f, c].Value = oo[x].cantidad;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                                workSheet.Cells[f, c].Value = oo[x].precio;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.0000"; f++;
                                workSheet.Cells[f, c].Value = oo[x].valor;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                break;
                            case "15020":
                                f = 17;
                                workSheet.Cells[f, c].Value = oo[x].cantidad;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                                workSheet.Cells[f, c].Value = oo[x].precio;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.0000"; f++;
                                workSheet.Cells[f, c].Value = oo[x].valor;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                break;

                        }
                        #endregion

                        

                    }

                }
                else
                {
                    workSheet.Cells[20, c].Value = "SIN FACTURA";
                }
                    
            }

            //excelPackage.Save();
            excelPackage.SaveAs(filesave);
            excelPackage = null;

        }


    }
}
