using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion.gestionATRGas
{
    public class Azucarera
    {
        logs.Log ficheroLog;
        utilidades.Param param;
        utilidades.ParamUser paramUser;
        sharePoint.Utilidades sharePoint;
        Dictionary<string, List<DateTime>> dic_excel_sin_datos;
        bool hay_error = true;
        sigame.SIGAME inventario_gas;
        utilidades.TelegramMensajes telegram;
        bool utiliza_telegram = false;


        public Azucarera()
        {
            inventario_gas = new sigame.SIGAME();
            dic_excel_sin_datos = new Dictionary<string, List<DateTime>>();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_Azucarera");
            param = new utilidades.Param("atrgas_param", servidores.MySQLDB.Esquemas.CON);
            paramUser = new utilidades.ParamUser("atrgas_param_user", System.Environment.UserName, servidores.MySQLDB.Esquemas.CON);
            sharePoint = new sharePoint.Utilidades();
            telegram = new utilidades.TelegramMensajes(param.GetValue("telegram_token_contratacion"), param.GetValue("telegram_channel_id"));
            utiliza_telegram = param.GetValue("utilizar_telegram") == "S";
        }

        public void Prueba()
        {
            BorrarContenidoDirectorio(param.GetValue("AZUCARERA_Carpeta_Descarga", DateTime.Now, DateTime.Now));
            DescargaArchivos();
        }

        public void Proceso()
        {

            try
            {
                if (HayQueLanzarProceso())                
                {
                    BorrarContenidoDirectorio(param.GetValue("AZUCARERA_Carpeta_Descarga", DateTime.Now, DateTime.Now));

                    if (!hay_error)
                        DescargaArchivos();

                    if (!hay_error)
                    {
                        GeneraCorreo(LeeExcel());
                        GeneraXML(LeeExcel());
                        Subir_a_FTP();
                        

                    }

                    if (!hay_error)
                        GeneraCorreoDatosFaltantes();

                    if (!hay_error)
                    {
                        ActualizaTablaMySQLProcesos();
                        Thread.Sleep(2000);                        
                        
                    }

                    if (!hay_error)
                        if (utiliza_telegram)
                            telegram.SendMessage("El proceso Azucarera ha finalizado correctamente.");


                }
            }catch(Exception e)
            {
                ficheroLog.AddError("Azucarera.Proceso: " + e.Message);
                if (utiliza_telegram)
                    telegram.SendMessage("CUIDADO ERROR!!!!"
                    + System.Environment.NewLine
                    + "================="
                    + "Necesario revisar proceso por ERROR: "
                    + System.Environment.NewLine
                    + e.Message);
            }
            
        }

        private void DescargaArchivos()
        {

            string destino = "";
            utilidades.Credenciales credenciales = new utilidades.Credenciales();
            office.Excel excel = new office.Excel();

            try
            {

                ficheroLog.Add("Inicio descarga de archivos Excel");
                ficheroLog.Add("================================");
                ficheroLog.Add("");

                sharePoint.userName = credenciales.server_user;
                sharePoint.password = credenciales.server_password;                     

                ficheroLog.Add("Conectando con el sitio: " + param.GetValue("AZUCARERA_siteURL", DateTime.Now, DateTime.Now));
                Console.WriteLine("Conectando con el sitio: " + param.GetValue("AZUCARERA_siteURL", DateTime.Now, DateTime.Now));

                sharePoint.siteURL = param.GetValue("AZUCARERA_siteURL", DateTime.Now, DateTime.Now);
                destino = param.GetValue("AZUCARERA_Carpeta_Descarga", DateTime.Now, DateTime.Now);

                if (param.ExisteParametroVigente("AZUCARERA_Excel_BAÑEZA", DateTime.Now))
                {
                    ficheroLog.Add("Descargando el archivo: " + param.GetValue("AZUCARERA_Excel_BAÑEZA", DateTime.Now, DateTime.Now));
                    Console.WriteLine("Descargando el archivo: " + param.GetValue("AZUCARERA_Excel_BAÑEZA", DateTime.Now, DateTime.Now));

                    excel.Abrir(param.GetValue("AZUCARERA_fileURL", DateTime.Now, DateTime.Now)
                        + param.GetValue("AZUCARERA_Excel_BAÑEZA", DateTime.Now, DateTime.Now));

                    excel.Save(destino + param.GetValue("AZUCARERA_Excel_BAÑEZA", DateTime.Now, DateTime.Now));

                    //sharePoint.fileURL = param.GetValue("AZUCARERA_fileURL", DateTime.Now, DateTime.Now)
                    //    + param.GetValue("AZUCARERA_Excel_BAÑEZA", DateTime.Now, DateTime.Now);
                    //sharePoint.destination = destino +
                    //    param.GetValue("AZUCARERA_Excel_BAÑEZA", DateTime.Now, DateTime.Now);
                    //sharePoint.Download();
                }

                if (param.ExisteParametroVigente("AZUCARERA_Excel_CHP", DateTime.Now))
                {
                    ficheroLog.Add("Descargando el archivo: " + param.GetValue("AZUCARERA_Excel_CHP", DateTime.Now, DateTime.Now));
                    Console.WriteLine("Descargando el archivo: " + param.GetValue("AZUCARERA_Excel_CHP", DateTime.Now, DateTime.Now));


                    excel.Abrir(param.GetValue("AZUCARERA_fileURL", DateTime.Now, DateTime.Now)
                       + param.GetValue("AZUCARERA_Excel_CHP", DateTime.Now, DateTime.Now));

                    excel.Save(destino + param.GetValue("AZUCARERA_Excel_CHP", DateTime.Now, DateTime.Now));

                    //sharePoint.fileURL = param.GetValue("AZUCARERA_fileURL", DateTime.Now, DateTime.Now)
                    //    + param.GetValue("AZUCARERA_Excel_CHP", DateTime.Now, DateTime.Now);
                    //sharePoint.destination = destino +
                    //    param.GetValue("AZUCARERA_Excel_CHP", DateTime.Now, DateTime.Now);
                    //sharePoint.Download();
                }



                if (param.ExisteParametroVigente("AZUCARERA_Excel_GUADALETE", DateTime.Now))
                {
                    ficheroLog.Add("Descargando el archivo: " + param.GetValue("AZUCARERA_Excel_GUADALETE", DateTime.Now, DateTime.Now));
                    Console.WriteLine("Descargando el archivo: " + param.GetValue("AZUCARERA_Excel_GUADALETE", DateTime.Now, DateTime.Now));


                    excel.Abrir(param.GetValue("AZUCARERA_fileURL", DateTime.Now, DateTime.Now)
                      + param.GetValue("AZUCARERA_Excel_GUADALETE", DateTime.Now, DateTime.Now));

                    excel.Save(destino + param.GetValue("AZUCARERA_Excel_GUADALETE", DateTime.Now, DateTime.Now));


                    //sharePoint.fileURL = param.GetValue("AZUCARERA_fileURL", DateTime.Now, DateTime.Now)
                    //    + param.GetValue("AZUCARERA_Excel_GUADALETE", DateTime.Now, DateTime.Now);
                    //sharePoint.destination = destino +
                    //    param.GetValue("AZUCARERA_Excel_GUADALETE", DateTime.Now, DateTime.Now);
                    //sharePoint.Download();
                }


                if (param.ExisteParametroVigente("AZUCARERA_Excel_MIRANDA", DateTime.Now))
                {
                    ficheroLog.Add("Descargando el archivo: " + param.GetValue("AZUCARERA_Excel_MIRANDA", DateTime.Now, DateTime.Now));
                    Console.WriteLine("Descargando el archivo: " + param.GetValue("AZUCARERA_Excel_MIRANDA", DateTime.Now, DateTime.Now));

                    excel.Abrir(param.GetValue("AZUCARERA_fileURL", DateTime.Now, DateTime.Now)
                     + param.GetValue("AZUCARERA_Excel_MIRANDA", DateTime.Now, DateTime.Now));

                    excel.Save(destino + param.GetValue("AZUCARERA_Excel_MIRANDA", DateTime.Now, DateTime.Now));


                    //sharePoint.fileURL = param.GetValue("AZUCARERA_fileURL", DateTime.Now, DateTime.Now)
                    //    + param.GetValue("AZUCARERA_Excel_MIRANDA", DateTime.Now, DateTime.Now);
                    //sharePoint.destination = destino +
                    //    param.GetValue("AZUCARERA_Excel_MIRANDA", DateTime.Now, DateTime.Now);
                    //sharePoint.Download();
                }


                if (param.ExisteParametroVigente("AZUCARERA_Excel_TORO", DateTime.Now))
                {
                    ficheroLog.Add("Descargando el archivo: " + param.GetValue("AZUCARERA_Excel_TORO", DateTime.Now, DateTime.Now));
                    Console.WriteLine("Descargando el archivo: " + param.GetValue("AZUCARERA_Excel_TORO", DateTime.Now, DateTime.Now));


                    excel.Abrir(param.GetValue("AZUCARERA_fileURL", DateTime.Now, DateTime.Now)
                     + param.GetValue("AZUCARERA_Excel_TORO", DateTime.Now, DateTime.Now));

                    excel.Save(destino + param.GetValue("AZUCARERA_Excel_TORO", DateTime.Now, DateTime.Now));

                    //sharePoint.fileURL = param.GetValue("AZUCARERA_fileURL", DateTime.Now, DateTime.Now)
                    //+ param.GetValue("AZUCARERA_Excel_TORO", DateTime.Now, DateTime.Now);
                    //sharePoint.destination = destino +
                    //    param.GetValue("AZUCARERA_Excel_TORO", DateTime.Now, DateTime.Now);
                    //sharePoint.Download();
                }


                ficheroLog.Add("Fin descarga de archivos Excel");
                ficheroLog.Add("================================");
                ficheroLog.Add("");
                hay_error = false;

            }
            catch(Exception e)
            {
                hay_error = true;
                if(utiliza_telegram)
                telegram.SendMessage("CUIDADO ERROR!!!!"
                   + System.Environment.NewLine
                   + "================="
                   + System.Environment.NewLine
                   + "Necesario revisar proceso Azucarera por ERROR en Azucarera.DescargaArchivo: "
                   + System.Environment.NewLine
                   + e.Message);
                ficheroLog.AddError("Azucarera.DescargaArchivo: " + e.Message);
            }
            


        }

        private List<EndesaEntity.contratacion.gas.Solicitud> LeeExcel()
        {
            List<EndesaEntity.contratacion.gas.Solicitud> listaSol =
                new List<EndesaEntity.contratacion.gas.Solicitud>();


            EndesaEntity.contratacion.gas.Solicitud sol;

            int c = 1;
            int f = 1;
            bool firstOnly = true;

            DateTime diaLectura = new DateTime();            
            DateTime fechaFutura = new DateTime();
            string cups = "";
            bool encontrado_dato_futuro = false;
            

            try
            {
                diaLectura = DateTime.Now.AddDays(1); // Hoy no Mañaaaaana
                fechaFutura = DateTime.Now.AddDays(Convert.ToInt32(param.GetValue("AZUCARERA_Dias_Excel_Chequeo", DateTime.Now, DateTime.Now))).AddDays(1); 


                ficheroLog.Add("Inicio lectura de archivos Excel");
                ficheroLog.Add("================================");
                ficheroLog.Add("");

                Console.WriteLine("");

                string[] listaArchivosExcel = Directory.GetFiles(param.GetValue("AZUCARERA_Carpeta_Descarga", DateTime.Now, DateTime.Now), "*.xlsx");
                for (int x = 0; x < listaArchivosExcel.Length; x++)
                {

                    FileInfo file = new FileInfo(listaArchivosExcel[x]);

                    ficheroLog.Add("Mirando dentro del archivo: " + file.Name);
                    Console.WriteLine("Mirando dentro del archivo: " + file.Name);

                    FileStream fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    ExcelPackage excelPackage = new ExcelPackage(fs);
                    var workSheet = excelPackage.Workbook.Worksheets.First();

                    sol = new EndesaEntity.contratacion.gas.Solicitud();
                    sol.cups = param.GetValue(file.Name, DateTime.Now, DateTime.Now);
                    sol.descripcion = file.Name.Replace("AZUCARERA_", "").Replace(".xlsx", "").ToUpper();

                    f = 1; // Porque la primera fila es la cabecera
                    for (int i = 0; i < 100000; i++)
                    {
                        f++;

                        if (Convert.ToDateTime(workSheet.Cells[f, 1].Value) == diaLectura.Date)
                        {
                            EndesaEntity.contratacion.gas.SolicitudDetalle d = new EndesaEntity.contratacion.gas.SolicitudDetalle();                            
                            d.fecha_inicio = Convert.ToDateTime(workSheet.Cells[f, 1].Value);
                            d.fecha_fin = Convert.ToDateTime(workSheet.Cells[f, 1].Value);
                            d.qd = Convert.ToDouble(workSheet.Cells[f, 2].Value) * 1000; // Leemos Mw pero guardamos kW
                            d.producto = "DIARIO";
                            sol.detalle.Add(d);
                            listaSol.Add(sol);
                            break;
                        }

                    }

                    for(DateTime d = diaLectura.Date; d <= fechaFutura; d = d.AddDays(1))
                    {
                        f = 1; // Porque la primera fila es la cabecera
                        encontrado_dato_futuro = false;
                        for (int i = 0; i < 100000; i++)
                        {
                            f++;

                            if (Convert.ToDateTime(workSheet.Cells[f, 1].Value) == d.Date)
                            {

                                if (workSheet.Cells[f, 2].Value != null)
                                {
                                    encontrado_dato_futuro = true;
                                    break;
                                }

                            }

                        }

                        if (!encontrado_dato_futuro)
                        {
                            List<DateTime> o;
                            if (!dic_excel_sin_datos.TryGetValue(file.Name, out o))
                            {
                                o = new List<DateTime>();
                                o.Add(d.Date);
                                dic_excel_sin_datos.Add(file.Name, o);
                            }
                            else
                                o.Add(d.Date);                                

                        }
                            
                    }


                    ficheroLog.Add("Cerrando el archivo: " + fs.Name);

                    fs.Close();
                    fs = null;
                    excelPackage.Dispose();
                    excelPackage = null;
                }

                ficheroLog.Add("Fin lectura de archivos Excel");
                ficheroLog.Add("================================");
                ficheroLog.Add("");

                hay_error = false;
                return listaSol;

            }
            catch (Exception e)
            {
                hay_error = true;
                ficheroLog.AddError("Azucarera.LeeExcel: " + e.Message);
                return null;
            }
        }

        private void GeneraCorreo(List<EndesaEntity.contratacion.gas.Solicitud> lista)
        {

                string body = "";            
                string subject = "";
                string from = "";
                string to = "";
                string cc = null;
                string attachment = null;
            
            try
            {

                if(lista.Count > 0)
                { 

                    //body = GeneraCuerpoHTML(lista);
                    from = param.GetValue("remitente", DateTime.Now, DateTime.Now);
                    //to = param.GetValue("AZUCARERA_Distribuidora_Gas", DateTime.Now, DateTime.Now);
                    cc = param.GetValue("remitente", DateTime.Now, DateTime.Now);
                    subject = param.GetValue("AZUCARERA_mail_asunto_distribuidora", DateTime.Now, DateTime.Now) + DateTime.Now.AddDays(1).ToString("yyyyMMdd");

                    attachment = System.Environment.CurrentDirectory + param.GetValue("mail_image_path", DateTime.Now, DateTime.Now);

                    EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                    //mes.SendMailWeb(from, to, cc, subject, body, attachment);


                    // Aviso al gestor
                    body = GeneraCuerpoHTMLGestores(lista);
                    to = param.GetValue("AZUCARERA_Mails_Aviso_Gestores", DateTime.Now, DateTime.Now);
                    subject = param.GetValue("AZUCARERA_mail_asunto_copia_gestor", DateTime.Now, DateTime.Now) + DateTime.Now.AddDays(1).ToString("yyyyMMdd");

                    string[] listaArchivosExcel = Directory.GetFiles(param.GetValue("AZUCARERA_Carpeta_Descarga", DateTime.Now, DateTime.Now), "*.xlsx");
                    for (int x = 0; x < listaArchivosExcel.Length; x++)
                    {
                        FileInfo file = new FileInfo(listaArchivosExcel[x]);
                        FileStream fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        
                        attachment += ";" + file.FullName;                            
                    }

                    
                    mes.SendMailWeb(from, to, cc, subject, body, attachment);
                    hay_error = false;

                }

            }
            catch(Exception e)
            {
                hay_error = true;
                ficheroLog.AddError("Azucarera.GeneraCorreo: " + e.Message);
            }
        }

        private void GeneraXML(List<EndesaEntity.contratacion.gas.Solicitud> lista)
        {
            cnmc.CNMC cnmc = new cnmc.CNMC();
            cnmc.XML formato_xml = new cnmc.XML();
            //EndesaBusiness.utilidades.ZIP zip = new utilidades.ZIP();
            contratacion.gestionATRGas.Distribuidoras distribuidoras = new Distribuidoras(true);
            int secuencial;
            string fileName = "";
            string fechaHora = "";
            string destinycompany = "";

            try
            {
                if(lista.Count() > 0)
                {
                    foreach (EndesaEntity.contratacion.gas.Solicitud p in lista)
                    {
                        secuencial = Convert.ToInt32(param.GetValue("secuencial_solicitud", DateTime.Now, DateTime.Now)) + 1;

                        EndesaEntity.cnmc.XML_A1_43 xml_a1_43 = new EndesaEntity.cnmc.XML_A1_43();

                        xml_a1_43.comreferencenum = param.GetValue("prefijo_solicitud", DateTime.Now, DateTime.Now)
                            + DateTime.Now.ToString("yyyy") + secuencial.ToString().PadLeft(4, '0');

                        xml_a1_43.destinycompany =
                            distribuidoras.Codigo_XML_CNMC_Distribuidora(param.GetValue("nombre_distribuidora_azucarera", DateTime.Now, DateTime.Now));
                        destinycompany = xml_a1_43.destinycompany;
                        xml_a1_43.dispatchingcompany = "0007";
                        xml_a1_43.documentnum = param.GetValue("AZUCARERA_NIF", DateTime.Now, DateTime.Now);
                        xml_a1_43.cups = p.cups;
                        xml_a1_43.productstartdate = p.detalle[0].fecha_inicio;
                        xml_a1_43.producttype = cnmc.Codigo_Tipo_Producto(p.detalle[0].producto);
                        xml_a1_43.producttolltype = cnmc.Codigo_Tipo_Peaje(
                            inventario_gas.Grupo_Presion(p.cups), p.detalle[0].qd * 330);
                        xml_a1_43.productqd = p.detalle[0].qd;
                        xml_a1_43.reqtype = 2;

                        // Estructura XML: A1_43_SCTD_0007_XXXX_yyyyMMdd_HHmmssfff.xml

                        fechaHora = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
                        fileName = param.GetValue("AZUCARERA_Carpeta_Descarga", DateTime.Now, DateTime.Now)
                            + param.GetValue("prefijo_archivo_a1_43", DateTime.Now, DateTime.Now)
                            + xml_a1_43.destinycompany + "_"
                            + fechaHora
                            + ".xml";
                        FileInfo file = new FileInfo(fileName);

                        formato_xml.CreaXML_A1_43(file, xml_a1_43);
                        GuardaNumSecuencialTemporal(secuencial);

                        for(int j = 0; j < p.detalle.Count; j++)
                        {
                            if (utiliza_telegram)
                                telegram.SendMessage("Solicitud para " + p.cups + " (" + p.descripcion + ")"
                              + System.Environment.NewLine
                              + "Producto: " + p.detalle[j].producto
                              + System.Environment.NewLine
                              + "Fecha Inicio: " + p.detalle[j].fecha_inicio.ToString("dd/MM/yyyy")
                              + System.Environment.NewLine
                              + "Fecha Fin: " + p.detalle[j].fecha_fin.ToString("dd/MM/yyyy")
                              + System.Environment.NewLine
                              + "Qd: " + p.detalle[j].qd.ToString("#.###") + " kWh");
                        }
                        

                    }

                    // Comprimimos los XML y los enviamos por mail
                    fechaHora = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    fileName = param.GetValue("AZUCARERA_Carpeta_Descarga", DateTime.Now, DateTime.Now)
                            + param.GetValue("prefijo_archivo_a1_43", DateTime.Now, DateTime.Now)
                            + destinycompany + "_"
                            + fechaHora
                            + ".zip";
                    FileInfo archivo = new FileInfo(fileName);
                    //zip.ComprimirVarios(param.GetValue("AZUCARERA_Carpeta_Descarga", DateTime.Now, DateTime.Now), 
                    //    ".*\\.(xml)$", archivo.FullName);

                    if (param.GetValue("enviar_mail_XML", DateTime.Now, DateTime.Now) == "S")
                        GeneraMail(fileName);
                }
                
            }
            catch (Exception e)
            {
                hay_error = true;
                ficheroLog.AddError("Azucarera.GeneraXML: " + e.Message);
            }
        }
        private void GeneraCorreoDatosFaltantes()
        {

            string body = "";
            string subject = "";
            string from = "";
            string to = "";
            string cc = null;
            string attachment = null;
            string lista_dias = "";
            bool firstOnly = true;

            DateTime fechaFutura = new DateTime();
                       

            try
            {
                fechaFutura = DateTime.Now.AddDays(Convert.ToInt32(param.GetValue("AZUCARERA_Dias_Excel_Chequeo", DateTime.Now, DateTime.Now)));

                if (dic_excel_sin_datos.Count > 0)
                {
                    // Aviso al gestor
                    body = (DateTime.Now.Hour > 14 ? "Buenas tardes." : "Buenos días.")
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "     Para los siguientes archivos no hay energía informada en las fechas:";
                        

                    foreach (KeyValuePair<string, List<DateTime>> p in dic_excel_sin_datos)                    
                    {

                        body += System.Environment.NewLine
                             + System.Environment.NewLine
                             + "    " + p.Key
                             + System.Environment.NewLine;

                        lista_dias = "";
                        for (int j = 0; j < p.Value.Count; j++)
                        {

                            body += "      - " + p.Value[j].ToString("dd/MM/yyyy")
                            + System.Environment.NewLine;
                        }

                        
                    }
                    
                    // Aviso al gestor

                    body += System.Environment.NewLine
                        + System.Environment.NewLine
                        + "Por favor, resuelvan lo antes posible."
                        + System.Environment.NewLine
                        + "Un saludo.";

                    from = param.GetValue("remitente", DateTime.Now, DateTime.Now);
                    to = param.GetValue("AZUCARERA_Mails_Aviso_Gestores", DateTime.Now, DateTime.Now);
                    subject = param.GetValue("AZUCARERA_Subject_Datos_Faltantes", DateTime.Now, DateTime.Now);
                    cc = param.GetValue("remitente", DateTime.Now, DateTime.Now);

                    EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                    mes.SendMail(from, to, cc, subject, body, attachment);

                }
                else
                {
                    ficheroLog.Add("Sin datos faltantes.");
                }

            }catch(Exception e)
            {
                hay_error = true;
                ficheroLog.AddError("GeneraCorreoDatosFaltantes: " + e.Message);
            }
        }


        public string GeneraCuerpoHTML(List<EndesaEntity.contratacion.gas.Solicitud> listaSol)
        {
            string body = "";
            string linea = "";
            
            try
            {               

                body = param.GetValue("html_body_head", DateTime.Now, DateTime.Now);
                for (int i = 0; i < listaSol.Count; i++)
                {
                    linea = param.GetValue("html_body_line", DateTime.Now, DateTime.Now);
                    linea = linea.Replace("CUPS", listaSol[i].cups);
                    // linea = linea.Replace("TARIFA_SOLICITADA", sol.detalle[i].tarifa);
                    linea = linea.Replace("CAUDAL_SOLICITADA", listaSol[i].detalle[0].qd.ToString());
                    linea = linea.Replace("FECHA_INICIO_SOLICITADA", listaSol[i].detalle[0].fecha_inicio.ToString("dd/MM/yyyy"));
                    linea = linea.Replace("FECHA_FIN_SOLICITADA", listaSol[i].detalle[0].fecha_fin.ToString("dd/MM/yyyy"));
                    linea = linea.Replace("PRODUCTO_SOLICITADO", listaSol[i].detalle[0].producto);
                    body += linea;
                }
                body += param.GetValue("html_body_foot", DateTime.Now, DateTime.Now);
                return body;
            }
            catch (Exception e)
            {
                hay_error = true;
                ficheroLog.AddError("GeneraCuerpoHTML: " + e.Message);
                return "";
            }
        }

        public string GeneraCuerpoHTMLGestores(List<EndesaEntity.contratacion.gas.Solicitud> listaSol)
        {
            string body = "";
            string linea = "";

            try
            {

                body = param.GetValue("html_body_head_gestor", DateTime.Now, DateTime.Now);
                for (int i = 0; i < listaSol.Count; i++)
                {
                    linea = param.GetValue("html_body_line_gestor", DateTime.Now, DateTime.Now);
                    linea = linea.Replace("CUPS", listaSol[i].cups);
                    linea = linea.Replace("PLANTA", listaSol[i].descripcion);                    
                    linea = linea.Replace("CAUDAL_SOLICITADA", listaSol[i].detalle[0].qd.ToString());
                    linea = linea.Replace("FECHA_INICIO_SOLICITADA", listaSol[i].detalle[0].fecha_inicio.ToString("dd/MM/yyyy"));
                    linea = linea.Replace("FECHA_FIN_SOLICITADA", listaSol[i].detalle[0].fecha_fin.ToString("dd/MM/yyyy"));
                    linea = linea.Replace("PRODUCTO_SOLICITADO", listaSol[i].detalle[0].producto);
                    body += linea;
                }
                body += param.GetValue("html_body_foot_gestor", DateTime.Now, DateTime.Now);
                return body;
            }
            catch (Exception e)
            {
                ficheroLog.AddError("GeneraCuerpoHTMLGestores: " + e.Message);
                hay_error = true;
                return "";
            }
        }

        private void BorrarContenidoDirectorio(string directorio)
        {
            string[] listaArchivos;
            FileInfo file;
            try
            {
                listaArchivos = Directory.GetFiles(directorio);
                for (int i = 0; i < listaArchivos.Count(); i++)
                {
                    file = new FileInfo(listaArchivos[i]);
                    file.Delete();
                }
                hay_error = false;
            }catch(Exception e)
            {
                hay_error = true;
                if (utiliza_telegram)
                    telegram.SendMessage("CUIDADO ERROR!!!!"
                    + System.Environment.NewLine
                    + "================="
                    + System.Environment.NewLine
                    + "Necesario revisar proceso Azucarera por ERROR en BorrarContenidoDirectorio: "
                    + System.Environment.NewLine
                    + e.Message);
                ficheroLog.AddError("Azucarera.BorrarContenidoDirectorio: " + e.Message);
            }
            
        }


        private void ActualizaTablaMySQLProcesos()
        {

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            DateTime ahora = new DateTime();

            // Actualizamos campos descripcion_autoconsumo de la tabla PS
            ahora = DateTime.Now;
            strSql = "update atrgas_fechas_procesos set fecha = '" + ahora.ToString("yyyy-MM-dd") + "'"
                + " where proceso = 'azucarera'";

            Console.WriteLine("Actualizamos la fecha de los procesos de la tabla en MySQL");
            
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        private bool HayQueLanzarProceso()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            DateTime hoy = new DateTime();

            try
            {

                hoy = DateTime.Now;
                strSql = "select fecha from atrgas_fechas_procesos where"
                   + " proceso = 'azucarera'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                if (r.Read())
                {
                    return (Convert.ToDateTime(r["fecha"]).Date < hoy.Date);
                }

                db.CloseConnection();
                return false;
            }
            catch(Exception e)
            {
                hay_error = true;
                return false;
            }

            



        }

        private void GuardaNumSecuencialTemporal(int secuencial_solicitud)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            strSql = "update atrgas_param set value = '" + secuencial_solicitud + "'"
                + " where code = 'secuencial_solicitud'";
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            // Volvemos a cargar parametros para tener el ultimo valor en memoria
            param = new utilidades.Param("atrgas_param", servidores.MySQLDB.Esquemas.CON);

        }

        private void GeneraMail(string adjunto)
        {
            string body = "";
            string subject = "";
            string from = "";
            string to = "";
            string cc = null;

            body = (DateTime.Now.Hour > 14 ? "Buenas tardes:" : "Buenos días:")
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "     Se adjunta ZIP con XML de Azucarera."
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "Un saludo.";

            from = param.GetValue("remitente", DateTime.Now, DateTime.Now);
            to = "cge@enel.com";
            subject = "Archivos XML Azucarera proceso SCTD A1 43";
            from = param.GetValue("remitente", DateTime.Now, DateTime.Now);

            EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
            mes.SendMail(from, to, cc, subject, body, adjunto);
        }

        public void Subir_a_FTP()
        {
            utilidades.UltimateFTP ftp;
            string ruta_destino = "";
            string distribuidora = "";

            try
            {
                string[] listaArchivos = Directory.GetFiles(param.GetValue("AZUCARERA_Carpeta_Descarga", DateTime.Now, DateTime.Now), "*.zip");
                if (listaArchivos.Length > 0)
                {

                    ficheroLog.Add("Conectando al FTP: " + param.GetValue("ftp_sctd_server", DateTime.Now, DateTime.Now));
                    ftp = new utilidades.UltimateFTP(
                        param.GetValue("ftp_sctd_server", DateTime.Now, DateTime.Now),
                        param.GetValue("ftp_sctd_user", DateTime.Now, DateTime.Now),
                        utilidades.FuncionesTexto.Decrypt(param.GetValue("ftp_sctd_pass", DateTime.Now, DateTime.Now), true),
                        param.GetValue("ftp_sctd_port", DateTime.Now, DateTime.Now));


                    for (int i = 0; i < listaArchivos.Length; i++)
                    {
                        FileInfo fichero = new FileInfo(listaArchivos[i]);
                        ficheroLog.Add("Detectado: " + fichero.Name);
                        distribuidora = fichero.Name.Substring(16, 4);
                        ruta_destino = param.GetValue("ftp_sctd_ruta_destino", DateTime.Now, DateTime.Now).Replace("distribuidora", distribuidora);
                        ftp.Upload(ruta_destino + fichero.Name, listaArchivos[i]);
                        if (utiliza_telegram)
                            telegram.SendMessage("Se publica en el FTP el archivo " + fichero.Name + " en la ruta " + ruta_destino);

                    }
                }
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Azucarera.Subir_a_FTP: " + e.Message);
                if (utiliza_telegram)
                    telegram.SendMessage("CUIDADO ERROR!!!!"
                   + System.Environment.NewLine
                   + "================="
                   + System.Environment.NewLine
                   + "Necesario revisar proceso por ERROR: "
                   + System.Environment.NewLine
                   + e.Message);
            }
        }
    }
}
