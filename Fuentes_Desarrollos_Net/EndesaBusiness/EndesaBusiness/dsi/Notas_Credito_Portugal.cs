using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EndesaBusiness.dsi
{
    public class Notas_Credito_Portugal
    {
        utilidades.Param p;
        logs.Log ficheroLog;
        string mensaje_correo = "";
        List<EndesaEntity.dsi.NotaCreditoPortugal> listaCreditos;
        public Notas_Credito_Portugal()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "DSI_NotasCreditoPortugal");
            p = new utilidades.Param("dsi_ncp_param", servidores.MySQLDB.Esquemas.FAC);
            listaCreditos = new List<EndesaEntity.dsi.NotaCreditoPortugal>();
        }

        public void Procesar()
        {
            
            utilidades.UltimateFTP ftp;
            bool renombrar = false;
            string ruta_inbox = "";
            string ruta_procesados = "";

            try
            {
                ruta_inbox = p.GetValue("inbox", DateTime.Now, DateTime.Now);
                ruta_procesados = p.GetValue("procesados", DateTime.Now, DateTime.Now);

                ficheroLog.Add("Inicio proceso");
                mensaje_correo = "Inicio proceso";

                ficheroLog.Add("Buscando archivos para subir al FTP en: " + ruta_inbox);
                Console.WriteLine("Buscando archivos para subir al FTP en: " + ruta_inbox);
                mensaje_correo += System.Environment.NewLine;
                mensaje_correo += "Buscando archivos para subir al FTP en: " + ruta_inbox;

                string[] listaArchivos = Directory.GetFiles(ruta_inbox, "*.xlsx");
                for(int i = 0; i < listaArchivos.Length; i++)
                {
                    ficheroLog.Add("Archivo encontrado: " + listaArchivos[i]);
                    Console.WriteLine("Archivo encontrado: " + listaArchivos[i]);
                }

                FileInfo fichero_hoy = new FileInfo(ruta_inbox + "NCR_BTN_D" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx");

                // Si tenemos más de un archivo
                // juntamos el contenido en el ultimo archivo
                // publicado.

                if(listaArchivos.Length > 1)
                    LeeExcel(listaArchivos);


                listaArchivos = Directory.GetFiles(ruta_inbox, "*.xlsx");

                renombrar = (!fichero_hoy.Exists);
                if (renombrar)
                    renombrar = listaArchivos.Length == 1;

                if (listaArchivos.Length > 0)
                {
                    ficheroLog.Add("Conectando al FTP: " + p.GetValue("ftp_server", DateTime.Now, DateTime.Now));
                    mensaje_correo += System.Environment.NewLine;
                    mensaje_correo += "Conectando al FTP: " + p.GetValue("ftp_server", DateTime.Now, DateTime.Now);
                                        
                    ftp = new utilidades.UltimateFTP(
                    p.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                    p.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                    utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_pass", DateTime.Now, DateTime.Now), true),
                    p.GetValue("ftp_port", DateTime.Now, DateTime.Now));
                    
                    

                    for (int i = 0; i < listaArchivos.Length; i++)
                    {
                        FileInfo fichero = new FileInfo(listaArchivos[i]);

                        if (renombrar)
                        {
                            fichero.CopyTo(fichero_hoy.FullName);
                            Thread.Sleep(1000);
                            fichero.Delete();

                            //fichero = fichero_hoy;

                            fichero = null;
                            fichero = new FileInfo(ruta_inbox + "NCR_BTN_D" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
                        }

                        ficheroLog.Add("Detectado: " + fichero.Name);
                        mensaje_correo += System.Environment.NewLine;
                        mensaje_correo += "Detectado: " + fichero.Name;

                        if (p.GetValue("utilizar_ftp") == "S")
                            ftp.Upload(p.GetValue("ruta_destino_FTP", DateTime.Now, DateTime.Now) + fichero.Name, listaArchivos[i]);


                        ficheroLog.Add("Moviendo: " + fichero.Name + " a " + ruta_procesados);
                        mensaje_correo += System.Environment.NewLine;
                        mensaje_correo += "Moviendo: " + fichero.Name + " a " + ruta_procesados;
                        fichero.MoveTo(ruta_procesados + fichero.Name);
                    }
                }
                else
                {
                    ficheroLog.Add("Sin ficheros para mover");
                    mensaje_correo += System.Environment.NewLine;
                    mensaje_correo += "Sin ficheros para mover";
                }

                ficheroLog.Add("Fin proceso");
                mensaje_correo += System.Environment.NewLine;
                mensaje_correo += "Fin proceso";

                EnvioMail();

            }
            catch(Exception e)
            {
                ficheroLog.AddError("Procesar: " + e.Message);
                EnvioMail_Error("Procesar: " + e.Message);

            }
        }

        private void EnvioMail()
        {
            string body = "";
            string subject = "";
            string from = "";
            string to = "";
            string cc = null;
            string attachment = null;

            from = p.GetValue("mail_remitente", DateTime.Now, DateTime.Now);
            to = p.GetValue("mail_destinatario", DateTime.Now, DateTime.Now);
            cc = p.GetValue("mail_cc", DateTime.Now, DateTime.Now);
            //body = p.GetValue("html_body", DateTime.Now, DateTime.Now);
            body = mensaje_correo;
            subject = p.GetValue("mail_asunto", DateTime.Now, DateTime.Now);
            
            EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
            mes.SendMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(body), attachment);
        }

        private void EnvioMail_Error(string mensaje)
        {
            string body = "";
            string subject = "";
            string from = "";
            string to = "";
            string cc = null;
            string attachment = null;

            from = p.GetValue("mail_remitente", DateTime.Now, DateTime.Now);
            to = p.GetValue("mail_supervisor", DateTime.Now, DateTime.Now);            
            
            body = mensaje;
            subject = "Error en proceso Notas_Credito_Portugal";
            
            EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
            mes.SendMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(body), attachment);
        }


        private void LeeExcel(string[] listaArchivos)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage;
            string linea = "";
            int c = 0;
            FileInfo fichero;
            int fila = 0;
            string ruta_procesados;
            try
            {
                ruta_procesados = p.GetValue("procesados", DateTime.Now, DateTime.Now);

                // Guardamos todos los datos del excel excepto del
                // ultimo excel donde vamos a añadir todas las lineas.

                for (int i = 0; i < listaArchivos.Length -1; i++)
                {
                    Thread.Sleep(1000);
                    fichero = new FileInfo(listaArchivos[i]);
                    excelPackage = new ExcelPackage(fichero);
                    var workSheet = excelPackage.Workbook.Worksheets.First();


                    for (int f = 2; f <= 60000; f++)
                    {
                        linea = "";

                        for (int columnas = 1; columnas <= 30; columnas++)
                            linea += workSheet.Cells[f, columnas].Value;

                        if(linea != "")
                        {
                            c = 1;
                            EndesaEntity.dsi.NotaCreditoPortugal p = new EndesaEntity.dsi.NotaCreditoPortugal();

                            if (workSheet.Cells[f, c].Value != null)
                                p.documento = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.empresa = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.idi = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.factura = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.importe = Convert.ToDouble(workSheet.Cells[f, c].Value.ToString()); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.tipo_factura = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.nif = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.contrato = workSheet.Cells[f, c].Value.ToString(); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                p.sec = workSheet.Cells[f, c].Value.ToString(); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                p.tipo = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.cliente = Convert.ToInt64(workSheet.Cells[f, c].Value.ToString()); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.atr = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.cups = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.potencia1 = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.rob = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.primfac = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.digi = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.ccaa = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.produc = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.costimp = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.costproc = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.costregul = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.lectreales = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.tifac = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.link = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.fecha = Convert.ToDateTime(workSheet.Cells[f, c].Value.ToString()); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.ctacomerc = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.primfaccch = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.factura2 = workSheet.Cells[f, c].Value.ToString(); c++;
                            if (workSheet.Cells[f, c].Value != null)
                                p.digital = workSheet.Cells[f, c].Value.ToString(); c++;

                            listaCreditos.Add(p);
                            
                        }
                        else
                        {
                            break;
                        }
                        
                    }
                    excelPackage = null;
                    FileInfo ficheroDestino = new FileInfo(ruta_procesados + fichero.Name);
                    if (!fichero.Exists)
                    {
                        fichero.MoveTo(ruta_procesados + fichero.Name);
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        fichero.MoveTo(ruta_procesados +
                         fichero.Name.Replace(".xlsx", "_" + DateTime.Now.ToString("HHmmss") + ".xlsx"));
                        Thread.Sleep(1000);
                    }                     

                }

                Thread.Sleep(1000);
                fichero = new FileInfo(listaArchivos[listaArchivos.Length -1]);
                excelPackage = new ExcelPackage(fichero);
                var workSheet2 = excelPackage.Workbook.Worksheets.First();

                // Buscamos la ultima fila para pegar el contenido del resto de archivos
                for (int f = 1; f <= 60000; f++)
                {
                    linea = "";
                    fila++;


                    for (int columnas = 1; columnas <= 30; columnas++)
                        linea += workSheet2.Cells[f, columnas].Value;
                    if (linea == "")
                        break;

                }

                fila--;
                foreach(EndesaEntity.dsi.NotaCreditoPortugal p in listaCreditos)
                {
                    fila++;
                    c = 1;
                    workSheet2.Cells[fila, c].Value = p.documento; c++;
                    workSheet2.Cells[fila, c].Value = p.empresa; c++;
                    workSheet2.Cells[fila, c].Value = p.idi; c++;
                    workSheet2.Cells[fila, c].Value = p.factura; c++;

                    workSheet2.Cells[fila, c].Value = p.importe;
                    workSheet2.Cells[fila, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    
                    workSheet2.Cells[fila, c].Value = p.tipo_factura; c++;
                    workSheet2.Cells[fila, c].Value = p.nif; c++;
                    workSheet2.Cells[fila, c].Value = p.contrato; c++;
                    workSheet2.Cells[fila, c].Value = p.sec; c++;
                    workSheet2.Cells[fila, c].Value = p.tipo; c++;

                    workSheet2.Cells[fila, c].Value = p.cliente; c++;
                    

                    workSheet2.Cells[fila, c].Value = p.atr; c++;
                    workSheet2.Cells[fila, c].Value = p.cups; c++;
                    workSheet2.Cells[fila, c].Value = p.potencia1; c++;
                    workSheet2.Cells[fila, c].Value = p.rob; c++;
                    workSheet2.Cells[fila, c].Value = p.primfac; c++;
                    workSheet2.Cells[fila, c].Value = p.digi; c++;
                    workSheet2.Cells[fila, c].Value = p.ccaa; c++;
                    workSheet2.Cells[fila, c].Value = p.produc; c++;
                    workSheet2.Cells[fila, c].Value = p.costimp; c++;
                    workSheet2.Cells[fila, c].Value = p.costproc; c++;
                    workSheet2.Cells[fila, c].Value = p.costregul; c++;
                    workSheet2.Cells[fila, c].Value = p.lectreales; c++;
                    workSheet2.Cells[fila, c].Value = p.tifac; c++;
                    workSheet2.Cells[fila, c].Value = p.link; c++;

                    workSheet2.Cells[fila, c].Value = p.fecha;
                    workSheet2.Cells[fila, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                    workSheet2.Cells[fila, c].Value = p.ctacomerc; c++;
                    workSheet2.Cells[fila, c].Value = p.primfaccch; c++;
                    workSheet2.Cells[fila, c].Value = p.factura2; c++;
                    workSheet2.Cells[fila, c].Value = p.digital; c++;                    

                }
                Thread.Sleep(1000);
                excelPackage.Save();
                excelPackage = null;

            }
            catch(Exception e)
            {
                ficheroLog.AddError("LeeExcel: " + e.Message);
                EnvioMail_Error("LeeExcel: " + e.Message);
            }
        }
    }
}
