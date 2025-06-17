using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.office
{
    class Excel
    {
        Microsoft.Office.Interop.Excel.Application xl;
        Microsoft.Office.Interop.Excel._Workbook libro;
        Microsoft.Office.Interop.Excel._Worksheet hoja;

        public enum Estilos
        {
            NEGRITA,
            ITALICA,
            SUBRAYADA
        }

        public Excel()
        {
            xl = new Microsoft.Office.Interop.Excel.Application();

            //Get a new workbook.
            libro = (Microsoft.Office.Interop.Excel._Workbook)(xl.Workbooks.Add(""));
            hoja = (Microsoft.Office.Interop.Excel._Worksheet)libro.ActiveSheet;
        }


        public Excel(string plantilla)
        {
            xl = new Microsoft.Office.Interop.Excel.Application();
            libro = xl.Workbooks.Open(plantilla);
            hoja = libro.Sheets[1];
        }
              



        public void Abrir(string archivo)
        {
            xl = new Microsoft.Office.Interop.Excel.Application();
            libro = xl.Workbooks.Open(archivo);
            hoja = libro.Sheets[1];
        }

        public void PonerTextoCeldaDerecha(int fila, int columna)
        {
            libro.ActiveSheet.Cells(fila, columna).HorizontalAlignment =
               Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
        }

        public void PonerTextoCeldaIzquierda(int fila, int columna)
        {
            libro.ActiveSheet.Cells(fila, columna).HorizontalAlignment =
               Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
        }

        public void PonerTextoCeldaCentrado(int fila, int columna)
        {
            libro.ActiveSheet.Cells(fila, columna).HorizontalAlignment =
               Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }

        public void DarFormatoFecha(int columna)
        {
            libro.ActiveSheet.Columns(columna).NumberFormat = "m/d/yyyy";
        }

        public void AlineaTextoDerecha(string columna)
        {
            libro.ActiveSheet.Range[columna].Style.HorizontalAlignment = HorizontalAlignment.Right;

        }

        public void CreaHoja(string nombreHoja)
        {
            int numHojas = libro.Sheets.Count;
            //libro.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            libro.Sheets.Add(After: libro.Sheets[libro.Sheets.Count]);

        }

        



        public void CambiaFormatoCelda(int fila, int columna)
        {
            libro.ActiveSheet.Cells(fila, columna).Style = "Comma";
            //libro.ActiveSheet.Cells(fila, columna).NumberFormat = @"_-* #,##0 _€_-;-* #,##0 _€_-;_-* ""-""?? _€_-;_-@_-";
        }

        public void AjustarAncho()
        {
            libro.ActiveSheet.Columns.AutoFit();
        }

        public void PonEstilo(int fila, int columna, Estilos estilo)
        {
            switch (estilo)
            {
                case Estilos.NEGRITA:
                    libro.ActiveSheet.Cells(fila, columna).Font.Bold = true;
                    break;
            }
        }

        public void EstiloMillares(int fila, int columna)
        {
            libro.ActiveSheet.Cells(fila, columna).NumberFormat = "#.#0,00";
        }


        public void PonValor(int fila, int columna, object valor)
        {
            libro.ActiveSheet.Cells(fila, columna).Value = valor;
        }

        public void Save(String nombreFichero)
        {

            libro.SaveAs(nombreFichero, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            libro.Close();
        }

        public void CambiaNombreHoja(string strNombreNuevo)
        {
            hoja.Name = strNombreNuevo;
        }

        public void SeleccionaHoja(int numeroHoja)
        {
            hoja = libro.Worksheets[numeroHoja];
            hoja.Select(Type.Missing);
        }

        public void Mostrar()
        {
            xl.Visible = true;
        }

        public void PonBorde(int fila, int columna)
        {
            Microsoft.Office.Interop.Excel.Range range = hoja.UsedRange;
            Microsoft.Office.Interop.Excel.Range cell = range.Cells[fila, columna];
            Microsoft.Office.Interop.Excel.Borders border = cell.Borders;

            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

        }

        public void CerrarSinGuardar()
        {
            libro.Application.ActiveWorkbook.Close(false);
        }

        public void Cerrar()
        {
            //xl.Quit();
            //libro = null;            
            Marshal.FinalReleaseComObject(hoja);
            libro.Close(Type.Missing, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(libro);

            xl.Quit();
            Marshal.FinalReleaseComObject(xl);


        }

        public void GuardarComo(string nombre, string formato)
        {
            // Excel.XlFileFormat.xlWorkbookNormal

            libro.Application.ActiveWorkbook.SaveAs(nombre, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
        }
    }
}
