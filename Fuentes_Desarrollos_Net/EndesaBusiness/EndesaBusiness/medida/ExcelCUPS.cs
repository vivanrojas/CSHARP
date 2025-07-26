using Microsoft.Exchange.WebServices.Data;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace EndesaBusiness.medida
{
    public class ExcelCUPS
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_ExcelCUPS");
        public bool hayError { get; set; }
        public string descripcion_error { get; set; }

        public List<EndesaEntity.medida.PuntoSuministro> lista_cups;
        public DateTime fecha_min { get; set; }
        public DateTime fecha_max { get; set; }

              

        public ExcelCUPS(string fichero)
        {
            fecha_min = new DateTime();
            fecha_max = new DateTime();
            fecha_min = DateTime.MaxValue;
            fecha_max = DateTime.MinValue;

            hayError = false;
            descripcion_error = "";
            lista_cups = new List<EndesaEntity.medida.PuntoSuministro>();
            CargaExcel(fichero);

        }

        public ExcelCUPS(string fichero, EmailMessage mail)
        {
            hayError = false;
            descripcion_error = "";
            lista_cups = new List<EndesaEntity.medida.PuntoSuministro>();
            CargaExcel(fichero, mail);
        }


        private void CargaExcel(string fichero, EmailMessage mail)
        {
            int c = 1;
            int f = 1;
            int id = 0;
            bool firstOnly = true;
            string cabecera = "";
            SolicitudCurvasGestoresFunciones sf = new SolicitudCurvasGestoresFunciones();
            SolicitudCurvasGestoresFuncionesDetalle sfd = new SolicitudCurvasGestoresFuncionesDetalle();

            try
            {
                // FileInfo file = new FileInfo(fichero);
                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();

                sf.mail = mail.From.Address;
                sf.fechahora_mail = mail.DateTimeReceived;

                lista_cups.Clear();

                f = 1; // Porque la primera fila es la cabecera
                for (int i = 0; i < 157; i++)
                {
                    c = 1;

                    if (firstOnly)
                    {
                        cabecera = workSheet.Cells[1, 1].Value.ToString()
                            + workSheet.Cells[1, 2].Value.ToString()
                            + workSheet.Cells[1, 3].Value.ToString();

                        if (!EstructuraCorrecta(cabecera))
                        {
                            this.hayError = true;
                            this.descripcion_error = "La estructura del archivo excel no es la correcta.";
                            sf.desc_error = "La estructura del archivo excel no es la correcta.";
                            sf.Save();
                            break;
                        }
                        else
                        {
                            sf.Save();
                        }
                        firstOnly = false;
                    }

                    f++;

                    if (workSheet.Cells[f, 1].Value == null
                        || workSheet.Cells[f, 2].Value == null
                        || workSheet.Cells[f, 3].Value == null)
                    {
                        break;
                    }
                    else
                    {

                        EndesaEntity.medida.PuntoSuministro cups = new EndesaEntity.medida.PuntoSuministro();
                        cups.id = id;
                        cups.cups20 = workSheet.Cells[f, c].Value.ToString().Trim().Substring(0, 20); c++;
                        cups.fd = Convert.ToDateTime(workSheet.Cells[f, c].Value.ToString().Trim()); c++;
                        cups.fh = Convert.ToDateTime(workSheet.Cells[f, c].Value.ToString().Trim());
                        if (cups.fd <= cups.fh && (cups.cups20.Length >= 20 && cups.cups20.Length <= 22))
                        {
                            id++;
                            lista_cups.Add(cups);

                            sfd.id = sf.id;
                            sfd.cups20 = cups.cups20;
                            sfd.fd = cups.fd;
                            sfd.fh = cups.fh;
                            sfd.Save();
                        }

                    }

                }

                fs.Close();
                fs = null;
                excelPackage = null;



            }
            catch (Exception e)
            {
                ficheroLog.AddError("CargaExcel: " + e.Message);
                this.hayError = true;
            }
        }


        // Dado un excel con las columnas CUPS20, FECHA DESDE y FECHA HASTA
        // devuelve una lista con los valores del Excel.
        private void CargaExcel(string fichero)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            bool firstOnly = true;
            string cabecera = "";

            try
            {
                // FileInfo file = new FileInfo(fichero);
                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();

                f = 1; // Porque la primera fila es la cabecera
                for (int i = 0; i < 5000; i++)
                {
                    c = 1;

                    if (firstOnly)
                    {
                        cabecera = workSheet.Cells[1, 1].Value.ToString()
                            + workSheet.Cells[1, 2].Value.ToString()
                            + workSheet.Cells[1, 3].Value.ToString();

                        if (!EstructuraCorrecta(cabecera))
                        {
                            this.hayError = true;
                            this.descripcion_error = "La estructura del archivo excel no es la correcta.";
                            break;
                        }

                        firstOnly = false;
                    }

                    f++;

                    if (workSheet.Cells[f, 1].Value == null)
                        break;

                    if (workSheet.Cells[f, 1].Value.ToString()
                        + workSheet.Cells[f, 2].Value.ToString()
                        + workSheet.Cells[f, 3].Value.ToString() == "")
                    {
                        break;
                    }
                    else
                    {
                        
                        EndesaEntity.medida.PuntoSuministro cups = new EndesaEntity.medida.PuntoSuministro();
                        cups.id = id;
                        cups.cups20 = workSheet.Cells[f, c].Value.ToString(); c++;
                        cups.fd = Convert.ToDateTime(workSheet.Cells[f, c].Value); c++;
                        cups.fh = Convert.ToDateTime(workSheet.Cells[f, c].Value);

                        fecha_min = fecha_min > cups.fd ? cups.fd : fecha_min;
                        fecha_max = fecha_max < cups.fh ? cups.fh : fecha_max;

                        lista_cups.Add(cups);
                        id++;
                    }

                }


                fs = null;
                excelPackage = null;

            }

            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                         "Error en el formato del fichero",
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Information);
            }
        }

        private bool EstructuraCorrecta(string cabecera)
        {


            if (cabecera.ToUpper().Trim() == "CUPS20FECHA DESDEFECHA HASTA")
                return true;
            else
                return false;
        }
    }
}
