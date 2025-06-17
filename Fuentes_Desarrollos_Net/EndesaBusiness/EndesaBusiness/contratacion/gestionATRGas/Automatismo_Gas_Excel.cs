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
using System.Xml;

namespace EndesaBusiness.contratacion.gestionATRGas
{
    public class Automatismo_Gas_Excel
    {
        logs.Log ficheroLog;
        utilidades.Param p;
        utilidades.TelegramMensajes telegram;
        bool hay_error = true;
        bool utiliza_telegram = false;
        Dictionary<string, List<DateTime>> dic_excel_sin_datos;

        EndesaBusiness.contratacion.gestionATRGas.GestionATRGas gestionATR;
        EndesaBusiness.contratacion.gestionATRGas.Distribuidoras distribuidoras =
            new EndesaBusiness.contratacion.gestionATRGas.Distribuidoras(true);
        EndesaBusiness.sigame.SIGAME inventario_gas;
        public bool error_ftp;
        public Automatismo_Gas_Excel()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_TomatesGuadiana");
            p = new utilidades.Param("atrgas_param", servidores.MySQLDB.Esquemas.CON);
            telegram = new utilidades.TelegramMensajes(p.GetValue("telegram_token_contratacion"), p.GetValue("telegram_channel_id"));
            gestionATR = new EndesaBusiness.contratacion.gestionATRGas.GestionATRGas();
            // utiliza_telegram = p.GetValue("utilizar_telegram") == "S";
        }


        public void Proceso()
        {

            try
            {
                if (HayQueLanzarProceso("tomates_guadiana_genera_solicitud"))
                {
                    BorrarContenidoDirectorio(p.GetValue("TOMATES_GUADIANA_Carpeta_Descarga_Excels"));

                    if (!hay_error)
                        DescargaArchivosExcel();

                    if (!hay_error)
                    {
                        List<EndesaEntity.contratacion.gas.Solicitud> lista_sol = LeeExcel();
                                                
                        GeneraXML(lista_sol);

                        GeneraCorreo(lista_sol);


                    }

                    //if (!hay_error)
                    //    GeneraCorreoDatosFaltantes();

                    if (!hay_error)
                    {
                        ActualizaTablaMySQLProcesos("tomates_guadiana_genera_solicitud");
                        Thread.Sleep(2000);

                    }

                    if (!hay_error)
                        if (utiliza_telegram)
                            telegram.SendMessage("El proceso TOMATES DEL GUADIANA S.COOP ha finalizado correctamente.");


                    
                }

                
            }
            catch (Exception e)
            {
                ficheroLog.AddError("TomatesGuadiana.Proceso: " + e.Message);
                if (utiliza_telegram)
                    telegram.SendMessage("CUIDADO ERROR!!!!"
                    + System.Environment.NewLine
                    + "================="
                    + System.Environment.NewLine
                    + "Necesario revisar proceso TomatesGuadiana por ERROR: "
                    + System.Environment.NewLine
                    + e.Message);
            }

        }

        private List<EndesaEntity.contratacion.gas.Solicitud> LeeExcel()
        {
            List<EndesaEntity.contratacion.gas.Solicitud> listaSol =
               new List<EndesaEntity.contratacion.gas.Solicitud>();

            EndesaBusiness.contratacion.gestionATRGas.ContratosGas contratosGas =
                new ContratosGas();


            EndesaEntity.contratacion.gas.Solicitud sol;

            int c = 1;
            int f = 1;
            bool firstOnly = true;

            DateTime diaLectura = new DateTime();
            DateTime fechaFutura = new DateTime();
            string cups = "";
            bool encontrado_dato_futuro = false;
            string codigo_producto = "";

            try
            {
                codigo_producto = GetLastActivationCode();
                if(codigo_producto != "")
                {
                    diaLectura = DateTime.Now.AddDays(1); // Hoy no Mañaaaaana               


                    ficheroLog.Add("Inicio lectura de archivos Excel");
                    ficheroLog.Add("================================");
                    ficheroLog.Add("");

                    Console.WriteLine("");

                    string[] listaArchivosExcel = Directory.GetFiles(p.GetValue("TOMATES_GUADIANA_Carpeta_Descarga_Excels", DateTime.Now, DateTime.Now), "*.xlsx");
                    for (int x = 0; x < listaArchivosExcel.Length; x++)
                    {

                        FileInfo file = new FileInfo(listaArchivosExcel[x]);

                        ficheroLog.Add("Mirando dentro del archivo: " + file.Name);
                        Console.WriteLine("Mirando dentro del archivo: " + file.Name);

                        FileStream fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                        ExcelPackage excelPackage = new ExcelPackage(fs);
                        var workSheet = excelPackage.Workbook.Worksheets.First();

                        sol = new EndesaEntity.contratacion.gas.Solicitud();
                        sol.cups = p.GetValue("TOMATES_GUADIANA_CUPS");
                        sol.descripcion = sol.cups.ToUpper();

                        f = 1; // Porque la primera fila es la cabecera
                        for (int i = 0; i < 100000; i++)
                        {
                            f++;

                            if (Convert.ToDateTime(workSheet.Cells[f, 1].Value) == diaLectura.Date)
                            {

                                if(Convert.ToString(workSheet.Cells[f, 2].Value) != "")
                                {
                                    EndesaEntity.contratacion.gas.SolicitudDetalle d = new EndesaEntity.contratacion.gas.SolicitudDetalle();
                                    d.fecha_inicio = Convert.ToDateTime(workSheet.Cells[f, 1].Value);
                                    d.fecha_fin = Convert.ToDateTime(workSheet.Cells[f, 1].Value);
                                    d.qd = Convert.ToDouble(workSheet.Cells[f, 2].Value); // Leemos kW

                                    d.producto = "DIARIO";

                                    d.codigo_producto = codigo_producto;

                                    if (GetLastQd() == d.qd)
                                        d.tipo_producto = 0;
                                    else
                                        d.tipo_producto = 1;

                                    sol.detalle.Add(d);
                                    listaSol.Add(sol);
                                    break;
                                }
                                else
                                {

                                    ficheroLog.Add("No se han encontrado datos en el Excel.");


                                    if (DateTime.Now.Hour > 17)
                                    {
                                        EnvioMail_DatosFaltantesExcel();
                                        hay_error = true;
                                        if (utiliza_telegram)
                                        {
                                            telegram.SendMessage("No se han encontrado datos en el Excel para el día "
                                            + DateTime.Now.AddDays(1).ToString("dd/MM/yyyy") + "."
                                            + System.Environment.NewLine
                                            + "El proceso no tiene más programaciones."
                                            + System.Environment.NewLine);
                                        }

                                        break;
                                    }
                                    else
                                    {
                                        if (utiliza_telegram)
                                        {
                                            telegram.SendMessage("No se han encontrado datos en el Excel para el día "
                                            + DateTime.Now.AddDays(1).ToString("dd/MM/yyyy") + "."
                                            + System.Environment.NewLine
                                            + "Se volverá a lanzar el proceso a las 19:03 horas."
                                            + System.Environment.NewLine);

                                            hay_error = true;
                                            break;
                                        }
                                            
                                    }

                                }


                                
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

                    if(listaSol.Count > 0)
                    {
                        hay_error = false;
                        return listaSol;
                    }
                    else
                    {
                        hay_error = true;
                        return null;
                    }

                   
                }
                else
                {
                    hay_error = true;
                    return null;
                }

                

            }
            catch (Exception e)
            {
                hay_error = true;
                ficheroLog.AddError("Automatismo_Gas_Excel.LeeExcel: " + e.Message);
                return null;
            }
        }

        private void GeneraXML(List<EndesaEntity.contratacion.gas.Solicitud> lista)
        {
            EndesaBusiness.cnmc.CNMC cnmc = new EndesaBusiness.cnmc.CNMC();
            EndesaBusiness.cnmc.XML formato_xml = new EndesaBusiness.cnmc.XML();
            EndesaBusiness.utilidades.ZipUnZip zip7z = new EndesaBusiness.utilidades.ZipUnZip();

            string secuencial;
            string fileName = "";
            string fechaHora = "";
            string destinycompany = "";

            bool firstOnly = true;
            string message_telegram = "";

            try
            {


                inventario_gas = new sigame.SIGAME();

                if (!Directory.Exists(p.GetValue("TOMATES_GUADIANA_inbox", DateTime.Now, DateTime.Now)))
                    Directory.CreateDirectory(p.GetValue("TOMATES_GUADIANA_inbox", DateTime.Now, DateTime.Now));

                BorrarContenidoDirectorio(p.GetValue("TOMATES_GUADIANA_inbox", DateTime.Now, DateTime.Now));

                if (lista.Count() > 0)
                {
                    foreach (EndesaEntity.contratacion.gas.Solicitud sol in lista)
                    {
                        for (int i = 0; i < sol.detalle.Count(); i++)
                        {

                            if (firstOnly)
                            {
                                if (p.GetValue("utilizar_telegram") == "S")
                                {
                                    message_telegram = "Solicitud para el cliente: "
                                    + System.Environment.NewLine
                                    + inventario_gas.NombreCliente(sol.cups)
                                    + System.Environment.NewLine
                                    + " del CUPS: " + sol.cups
                                    + System.Environment.NewLine
                                    + "Para la distribuidora: "
                                    + System.Environment.NewLine
                                    + inventario_gas.Distribuidora(sol.cups)
                                    + System.Environment.NewLine;

                                    
                                }

                                ficheroLog.Add("Solicitud para el cliente: " + inventario_gas.NombreCliente(sol.cups));
                                ficheroLog.Add("Para el CUPS: " + sol.cups);
                                ficheroLog.Add("Para la distribuidora: " + inventario_gas.Distribuidora(sol.cups));


                                firstOnly = false;
                            }


                            Thread.Sleep(1000);
                            secuencial = DateTime.Now.ToString("yyMMddHHmmss").ToString();
                            gestionATR.GuardaNumSecuencialTemporal(secuencial);
                            Thread.Sleep(1000);

                            EndesaEntity.cnmc.XML_A1_43 xml_a1_43 = new EndesaEntity.cnmc.XML_A1_43();                            

                            xml_a1_43.comreferencenum = secuencial;

                            xml_a1_43.destinycompany =
                                    distribuidoras.Codigo_XML_CNMC_Distribuidora(inventario_gas.Distribuidora(sol.cups).ToUpper());
                            destinycompany = xml_a1_43.destinycompany;
                            xml_a1_43.dispatchingcompany = "0007";
                            xml_a1_43.documentnum = inventario_gas.NIF(sol.cups);
                            xml_a1_43.cups = sol.cups;
                            xml_a1_43.productstartdate = sol.detalle[i].fecha_inicio;
                            xml_a1_43.reqtype = sol.detalle[i].tipo_producto;
                            xml_a1_43.productcode = sol.detalle[i].codigo_producto;
                            xml_a1_43.producttype = cnmc.Codigo_Tipo_Producto(sol.detalle[i].producto.ToUpper());                            
                            xml_a1_43.producttolltype = cnmc.Codigo_Tipo_Peaje(inventario_gas.Tarifa(sol.cups));

                            xml_a1_43.productqd = sol.detalle[i].qd;
                            xml_a1_43.productqi = sol.detalle[i].qi;
                            xml_a1_43.startHour = sol.detalle[i].hora_inicio;



                            if (p.GetValue("utilizar_telegram") == "S")
                            {
                                
                                message_telegram += System.Environment.NewLine
                                 + "Producto: " + sol.detalle[i].producto
                                 + System.Environment.NewLine
                                 + "Fecha Inicio: " + sol.detalle[i].fecha_inicio.ToString("dd/MM/yyyy")
                                 + System.Environment.NewLine
                                 + "Fecha Fin: " + sol.detalle[i].fecha_fin.ToString("dd/MM/yyyy")
                                 + System.Environment.NewLine
                                 + "Qd: " + String.Format("{0:N0}", sol.detalle[i].qd) + " kWh";
                                
                            }

                            ficheroLog.Add("Producto: " + sol.detalle[i].producto);
                            ficheroLog.Add("Fecha Inicio: " + sol.detalle[i].fecha_inicio.ToString("dd/MM/yyyy"));
                            ficheroLog.Add("Fecha Fin: " + sol.detalle[i].fecha_fin.ToString("dd/MM/yyyy"));
                            ficheroLog.Add("Qd " + String.Format("{0:N0}", sol.detalle[i].qd) + " kWh");

                            fechaHora = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
                            fileName = p.GetValue("TOMATES_GUADIANA_inbox", DateTime.Now, DateTime.Now)
                                + @"\" + p.GetValue("prefijo_archivo_a1_43", DateTime.Now, DateTime.Now)
                                + xml_a1_43.destinycompany + "_"
                                + fechaHora
                                + ".xml";
                            FileInfo file = new FileInfo(fileName);

                            formato_xml.CreaXML_A1_43(file, xml_a1_43);

                            file.CopyTo(p.GetValue("TOMATES_GUADIANA_history") + file.Name, true);

                        }

                        telegram.SendMessage(message_telegram);

                        //fechaHora = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        //fileName = p.GetValue("TOMATES_GUADIANA_inbox", DateTime.Now, DateTime.Now) + @"\"
                        //        + p.GetValue("prefijo_archivo_a1_43", DateTime.Now, DateTime.Now)
                        //        + destinycompany + "_"
                        //        + fechaHora
                        //        + ".zip";
                        //FileInfo archivo = new FileInfo(fileName);

                        //zip7z.ComprimirVarios(p.GetValue("TOMATES_GUADIANA_inbox", DateTime.Now, DateTime.Now) + "\\" + "*.xml",
                        //    archivo.FullName); 

                        //if (sol.distribuidora_tramitacion == "XML_SCTD")
                        //    gestionATR.Subir_a_FTP();
                        //else if (sol.distribuidora_tramitacion == "XML")
                        //    gestionATR.Subir_a_FTP_Extremadura();

                        Subir_a_FTP_Extremadura();

                    }
                }


            }
            catch (Exception e)
            {
                if (utiliza_telegram)
                    telegram.SendMessage("CUIDADO ERROR!!!!"
                    + System.Environment.NewLine
                    + "================="
                    + System.Environment.NewLine
                    + "Necesario revisar proceso Tomates del Guadiana por ERROR en GeneraXML: "
                    + System.Environment.NewLine
                    + e.Message);
                ficheroLog.AddError("TOMATES DEL GUADIANA S.COOP.GeneraXML: " + e.Message);
            }
        }

        public void DescargaDatosFTPActivaciones()
        {
            EndesaBusiness.utilidades.UltimateFTP ftp;
            DateTime ultimoDiaLectura = new DateTime();
            string[] listaArchivosXML;
            List<EndesaEntity.contratacion.gas.A2_43> listaXML = new List<EndesaEntity.contratacion.gas.A2_43>();

            try
            {
                ficheroLog.Add("Comprobamos si hay que lanzar el proceso: " +
                    HayQueLanzarProceso("tomates_guadiana_lectura_activaciones"));
                if (HayQueLanzarProceso("tomates_guadiana_lectura_activaciones"))
                {

                    ultimoDiaLectura = DateTime.Now.Date;

                    for (DateTime d = ultimoDiaLectura.Date; d <= DateTime.Now.Date; d = d.AddDays(1))
                    {
                        string servidor = p.GetValue("ftp_extremadura_server_activacion");
                        servidor = servidor.Replace("YYYYMMDD", d.ToString("yyyyMMdd"));

                        ficheroLog.Add("Obteniendo datos de la ruta: " + servidor);

                        ficheroLog.Add("Conectando al FTP: " + servidor);
                        ftp = new EndesaBusiness.utilidades.UltimateFTP(servidor,
                            p.GetValue("ftp_extremadura_user"),
                            utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_extremadura_pass"), true),
                            p.GetValue("ftp_extremadura_port"));

                        ficheroLog.Add("Descargando datos en " + p.GetValue("inbox_activacion"));
                        ftp.DownloadAllInSecureFTP(p.GetValue("inbox_activacion"));

                    }

                    listaArchivosXML = Directory.GetFiles(p.GetValue("inbox_activacion"), "*.XML");
                    for (int j = 0; j < listaArchivosXML.Length; j++)
                    {

                        FileInfo ficheroXML = new FileInfo(listaArchivosXML[j]);
                        ficheroLog.Add("Leyendo archivo " + ficheroXML.FullName);
                        EndesaEntity.contratacion.gas.A2_43 xml = CargaXLM(ficheroXML.FullName);
                        if(xml.cups == p.GetValue("TOMATES_GUADIANA_CUPS") && xml.productcode != null)
                        {
                            listaXML.Add(xml);


                            // Si ya existe el archivo lo borramos
                            FileInfo ficheroXMLDestino = 
                                new FileInfo(p.GetValue("inbox_activacion") + @"\history\" + ficheroXML.Name);

                            if (ficheroXML.Exists)
                                ficheroXMLDestino.Delete();

                            ficheroXML.CopyTo(p.GetValue("inbox_activacion") + @"\history\" + ficheroXML.Name);
                        }
                        
                        ficheroXML.Delete();
                    }

                    if (listaXML.Count > 0)
                    {
                        ficheroLog.Add("Guardando XMLs en BBDD ");
                        GuardadoBBDD(listaXML);
                        ActualizaTablaMySQLProcesos("tomates_guadiana_lectura_activaciones");
                    }

                    
                }

            }
            catch(Exception ex)
            {
                ficheroLog.AddError("DescargaDatosFTPActivaciones " + ex.Message);
                if (p.GetValue("utilizar_telegram") == "S")
                    telegram.SendMessage("CUIDADO ERROR!!!!"
                    + System.Environment.NewLine
                    + "================="
                    + System.Environment.NewLine
                    + "Necesario revisar proceso DescargaDatosFTPActivaciones TomatesGuadiana por ERROR: "
                    + System.Environment.NewLine
                    + ex.Message);
            }
        }
        public void DescargaDatosFTPAceptaciones()
        {
            EndesaBusiness.utilidades.UltimateFTP ftp;
            DateTime ultimoDiaLectura = new DateTime();
            string[] listaArchivosXML;
            List<EndesaEntity.contratacion.gas.A2_43> listaXML = new List<EndesaEntity.contratacion.gas.A2_43>();

            try
            {
                ficheroLog.Add("Comprobamos si hay que lanzar el proceso: " +
                    HayQueLanzarProceso("tomates_guadiana_lectura_aceptaciones"));
                if (HayQueLanzarProceso("tomates_guadiana_lectura_aceptaciones"))
                {

                    ultimoDiaLectura = UltimoDiaAcepetacion().AddDays(1);

                    for (DateTime d = ultimoDiaLectura.Date; d <= DateTime.Now.Date; d = d.AddDays(1))
                    {
                        string servidor = p.GetValue("ftp_extremadura_server_aceptacion");
                        servidor = servidor.Replace("YYYYMMDD", d.ToString("yyyyMMdd"));

                        ficheroLog.Add("Obteniendo datos de la ruta: " + servidor);

                        ficheroLog.Add("Conectando al FTP: " + servidor);
                        ftp = new EndesaBusiness.utilidades.UltimateFTP(servidor,
                            p.GetValue("ftp_extremadura_user"),
                            utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_extremadura_pass"), true),
                            p.GetValue("ftp_extremadura_port"));

                        ficheroLog.Add("Descargando datos en " + p.GetValue("inbox_aceptacion"));
                        ftp.DownloadAllInSecureFTP(p.GetValue("inbox_aceptacion"));

                    }

                    listaArchivosXML = Directory.GetFiles(p.GetValue("inbox_aceptacion"), "*.XML");
                    for (int j = 0; j < listaArchivosXML.Length; j++)
                    {

                        FileInfo ficheroXML = new FileInfo(listaArchivosXML[j]);
                        ficheroLog.Add("Leyendo archivo " + ficheroXML.FullName);
                        EndesaEntity.contratacion.gas.A2_43 xml = CargaXLM(ficheroXML.FullName);
                        if (xml.cups == p.GetValue("TOMATES_GUADIANA_CUPS"))
                        {
                            listaXML.Add(xml);
                            ficheroXML.CopyTo(p.GetValue("inbox_aceptacion") + @"\history\" + ficheroXML.Name);
                        }

                        ficheroXML.Delete();
                    }

                    if (listaXML.Count > 0)
                    {
                        ficheroLog.Add("Guardando XMLs en BBDD ");
                        GuardadoBBDD_Aceptaciones(listaXML);
                        ActualizaTablaMySQLProcesos("tomates_guadiana_lectura_aceptaciones");
                    }

                    
                }

            }
            catch (Exception ex)
            {
                ficheroLog.AddError("DescargaDatosFTPAceptaciones " + ex.Message);
                if (p.GetValue("utilizar_telegram") == "S")
                    telegram.SendMessage("CUIDADO ERROR!!!!"
                    + System.Environment.NewLine
                    + "================="
                    + System.Environment.NewLine
                    + "Necesario revisar proceso DescargaDatosFTPAceptaciones TomatesGuadiana por ERROR: "
                    + System.Environment.NewLine
                    + ex.Message);
            }
        }

        private EndesaEntity.contratacion.gas.A2_43 CargaXLM(string fileName)
        {
            string cod_ini = "";
            string cod_fin = "";
            string valor = "";

            EndesaEntity.contratacion.gas.A2_43 xml = new EndesaEntity.contratacion.gas.A2_43();

            try
            {
                FileInfo file = new FileInfo(fileName);
                xml.fichero = file.Name;
                XmlTextReader r = new XmlTextReader(fileName);
                while (r.Read())
                {
                    switch (r.NodeType)
                    {
                        case XmlNodeType.Element: // The node is an element.
                            cod_ini = r.Name;
                            break;

                        case XmlNodeType.Text: //Display the text in each element.
                            valor = EndesaBusiness.utilidades.FuncionesTexto.ArreglaAcentos(r.Value);
                            break;
                        case XmlNodeType.EndElement: //Display the end of the element.
                            cod_fin = r.Name;
                            break;
                    }

                    if (cod_ini == cod_fin)
                        switch (cod_ini)
                        {
                            case "dispatchingcode":
                                xml.dispatchingcode = valor;
                                break;
                            case "dispatchingcompany":
                                xml.dispatchingcompany = valor;
                                break;
                            case "destinycompany":
                                xml.destinycompany = valor;
                                break;
                            case "communicationsdate":
                                xml.communicationsdate = Convert.ToDateTime(valor); 
                                break;
                            case "communicationshour":
                                xml.communicationshour = valor;
                                break;
                            case "processcode":
                                xml.processcode = valor;
                                break;
                            case "messagetype":
                                xml.messagetype = valor;
                                break;
                            case "comreferencenum":
                                xml.comreferencenum = valor;
                                break;
                            case "responsedate":
                                xml.responsedate = Convert.ToDateTime(valor); 
                                break;
                            case "responsehour":
                                xml.responsehour = valor;
                                break;
                            case "result":
                                xml.result = valor;
                                break;
                            case "resultdesc":
                                xml.resultdesc = valor;
                                break;
                            case "nationality":
                                xml.nationality = valor;
                                break;
                            case "documenttype":
                                xml.documenttype = valor;
                                break;
                            case "documentnum":
                                xml.documentnum = valor;
                                break;
                            case "cups":
                                xml.cups = valor;
                                break;
                            case "productcode":
                                xml.productcode = valor;
                                break;
                            case "producttype":
                                xml.producttype = valor;
                                break;
                            case "producttolltype":
                                xml.producttolltype = valor;
                                break;
                            case "productqd":
                                xml.productqd = Convert.ToInt32(valor);
                                break;
                            case "productqa":
                                xml.productqa = Convert.ToInt32(valor);
                                break;
                            case "productenddate":
                                xml.productenddate = Convert.ToDateTime(valor); 
                                break;
                            case "resultreason":
                                xml.resultreason = valor;
                                break;
                            case "resultreasondesc":
                                xml.resultreasondesc = valor;
                                break;
                            case "extrainfo":
                                xml.extrainfo = valor;
                                break;

                        }


                }
                r.Close();
                return xml;
            }
            catch(Exception ex)
            {
                ficheroLog.AddError("CargaXLM " + ex.Message);
                return null;
            }
        }


        private void GuardadoBBDD(List<EndesaEntity.contratacion.gas.A2_43> lista)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            bool firstOnly = true;
            int num_reg = 0;

            foreach (EndesaEntity.contratacion.gas.A2_43 xml in lista)
            {
                if (firstOnly)
                {
                    strSql = "replace into atrgas_activaciones (dispatchingcode, dispatchingcompany, destinycompany,"
                        + " communicationsdate, communicationshour, processcode, messagetype, comreferencenum, responsedate, responsehour,"
                        + " result, resultdesc, nationality, documenttype, documentnum, cups, productcode,"
                        + " producttype, producttolltype, productqd, productqa, productenddate, file, created_by"
                        + " ) values ";
                    firstOnly = false;
                }

                num_reg++;

                #region Campos

                if (xml.dispatchingcode != null)
                    strSql += "('" + xml.dispatchingcode + "'";
                else
                    strSql += "(null";

                if (xml.dispatchingcompany != null)
                    strSql += ", '" + xml.dispatchingcompany + "'";
                else
                    strSql += ", null";

                if (xml.destinycompany != null)
                    strSql += ", '" + xml.destinycompany + "'";
                else
                    strSql += ", null";

                if (xml.communicationsdate != null)
                    strSql += ", '" + xml.communicationsdate.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ", null";

                if (xml.communicationshour != null)
                    strSql += ", '" + xml.communicationshour + "'";
                else
                    strSql += ", null";

                if (xml.processcode != null)
                    strSql += ", '" + xml.processcode + "'";
                else
                    strSql += ", null";

                if (xml.messagetype != null)
                    strSql += ", '" + xml.messagetype + "'";
                else
                    strSql += ", null";               

                if (xml.comreferencenum != null)
                    strSql += ", '" + xml.comreferencenum + "'";
                else
                    strSql += ", null";

                if (xml.responsedate != null)
                    strSql += ", '" + xml.responsedate.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ", null";

                if (xml.responsehour != null)
                    strSql += ", '" + xml.responsehour + "'";
                else
                    strSql += ", null";

                if (xml.result != null)
                    strSql += ", '" + xml.result + "'";
                else
                    strSql += ", null";

                if (xml.resultdesc != null)
                    strSql += ", '" + xml.resultdesc + "'";
                else
                    strSql += ", null";

                if (xml.nationality != null)
                    strSql += ", '" + xml.nationality + "'";
                else
                    strSql += ", null";

                if (xml.documenttype != null)
                    strSql += ", '" + xml.documenttype + "'";
                else
                    strSql += ", null";

                if (xml.documentnum != null)
                    strSql += ", '" + xml.documentnum + "'";
                else
                    strSql += ", null";

                if (xml.cups != null)
                    strSql += ", '" + xml.cups + "'";
                else
                    strSql += ", null";

                if (xml.productcode != null)
                    strSql += ", '" + xml.productcode + "'";
                else
                    strSql += ", null";

                if (xml.producttype != null)
                    strSql += ", '" + xml.producttype + "'";
                else
                    strSql += ", null";

                if (xml.producttolltype != null)
                    strSql += ", '" + xml.producttolltype + "'";
                else
                    strSql += ", null";

                if (xml.productqd != 0)
                    strSql += ", " + xml.productqd;
                else
                    strSql += ", null";

                if (xml.productqa != 0)
                    strSql += ", " + xml.productqa;
                else
                    strSql += ", null";

                if (xml.productenddate != null)
                    strSql += ", '" + xml.productenddate.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ", null";

                if (xml.fichero != null)
                    strSql += ", '" + xml.fichero + "'";
                else
                    strSql += ", null";                            

                strSql += ", '" + System.Environment.UserName + "'),";
                #endregion

                if (num_reg > 250)
                {
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    num_reg = 0;
                    strSql = "";
                    firstOnly = true;
                }

            }

            if (num_reg > 0)
            {
                db = new EndesaBusiness.servidores.MySQLDB(EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                num_reg = 0;
                strSql = "";
            }


        }

        private void GuardadoBBDD_Aceptaciones(List<EndesaEntity.contratacion.gas.A2_43> lista)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            bool firstOnly = true;
            int num_reg = 0;

            foreach (EndesaEntity.contratacion.gas.A2_43 xml in lista)
            {
                if (firstOnly)
                {
                    strSql = "replace into atrgas_aceptaciones (dispatchingcode, dispatchingcompany, destinycompany,"
                        + " communicationsdate, communicationshour, processcode, messagetype, comreferencenum, responsedate, responsehour,"
                        + " result, resultdesc, resultreason, resultreasondesc, nationality, documenttype, documentnum, cups, productcode,"
                        + " producttype, producttolltype, productqd, productqa, productenddate, extrainfo, file, created_by"
                        + " ) values ";
                    firstOnly = false;
                }

                num_reg++;

                #region Campos

                if (xml.dispatchingcode != null)
                    strSql += "('" + xml.dispatchingcode + "'";
                else
                    strSql += "(null";

                if (xml.dispatchingcompany != null)
                    strSql += ", '" + xml.dispatchingcompany + "'";
                else
                    strSql += ", null";

                if (xml.destinycompany != null)
                    strSql += ", '" + xml.destinycompany + "'";
                else
                    strSql += ", null";

                if (xml.communicationsdate != null)
                    strSql += ", '" + xml.communicationsdate.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ", null";

                if (xml.communicationshour != null)
                    strSql += ", '" + xml.communicationshour + "'";
                else
                    strSql += ", null";

                if (xml.processcode != null)
                    strSql += ", '" + xml.processcode + "'";
                else
                    strSql += ", null";

                if (xml.messagetype != null)
                    strSql += ", '" + xml.messagetype + "'";
                else
                    strSql += ", null";

                if (xml.comreferencenum != null)
                    strSql += ", '" + xml.comreferencenum + "'";
                else
                    strSql += ", null";

                if (xml.responsedate != null)
                    strSql += ", '" + xml.responsedate.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ", null";

                if (xml.responsehour != null)
                    strSql += ", '" + xml.responsehour + "'";
                else
                    strSql += ", null";

                if (xml.result != null)
                    strSql += ", '" + xml.result + "'";
                else
                    strSql += ", null";

                if (xml.resultdesc != null)
                    strSql += ", '" + xml.resultdesc + "'";
                else
                    strSql += ", null";

                if (xml.resultreason != null)
                    strSql += ", '" + xml.resultreason + "'";
                else
                    strSql += ", null";

                if (xml.resultreasondesc != null)
                    strSql += ", '" + xml.resultreasondesc + "'";
                else
                    strSql += ", null";

                if (xml.nationality != null)
                    strSql += ", '" + xml.nationality + "'";
                else
                    strSql += ", null";

                if (xml.documenttype != null)
                    strSql += ", '" + xml.documenttype + "'";
                else
                    strSql += ", null";

                if (xml.documentnum != null)
                    strSql += ", '" + xml.documentnum + "'";
                else
                    strSql += ", null";

                if (xml.cups != null)
                    strSql += ", '" + xml.cups + "'";
                else
                    strSql += ", null";

                if (xml.productcode != null)
                    strSql += ", '" + xml.productcode + "'";
                else
                    strSql += ", null";

                if (xml.producttype != null)
                    strSql += ", '" + xml.producttype + "'";
                else
                    strSql += ", null";

                if (xml.producttolltype != null)
                    strSql += ", '" + xml.producttolltype + "'";
                else
                    strSql += ", null";

                if (xml.productqd != 0)
                    strSql += ", " + xml.productqd;
                else
                    strSql += ", null";

                if (xml.productqa != 0)
                    strSql += ", " + xml.productqa;
                else
                    strSql += ", null";

                if (xml.productenddate != null)
                    strSql += ", '" + xml.productenddate.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ", null";

                if (xml.extrainfo != null)
                    strSql += ", '" + xml.extrainfo + "'";
                else
                    strSql += ", null";

                if (xml.fichero != null)
                    strSql += ", '" + xml.fichero + "'";
                else
                    strSql += ", null";

                strSql += ", '" + System.Environment.UserName + "'),";
                #endregion

                if (num_reg > 250)
                {
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    num_reg = 0;
                    strSql = "";
                    firstOnly = true;
                }

            }

            if (num_reg > 0)
            {
                db = new EndesaBusiness.servidores.MySQLDB(EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                num_reg = 0;
                strSql = "";
            }


        }

        private void DescargaArchivosExcel()
        {

            string destino = "";
            utilidades.Credenciales credenciales = new utilidades.Credenciales();
            office.Excel excel = new office.Excel();
            sharePoint.Utilidades sharePoint;

            try
            {

                sharePoint = new sharePoint.Utilidades();

                ficheroLog.Add("Inicio descarga de archivos Excel");
                ficheroLog.Add("================================");
                ficheroLog.Add("");

                sharePoint.userName = credenciales.server_user;
                sharePoint.password = credenciales.server_password;

                ficheroLog.Add("Conectando con el sitio: " + p.GetValue("TOMATES_GUADIANA_siteURL"));
                Console.WriteLine("Conectando con el sitio: " + p.GetValue("TOMATES_GUADIANA_siteURL"));

                sharePoint.siteURL = p.GetValue("TOMATES_GUADIANA_siteURL");
                destino = p.GetValue("TOMATES_GUADIANA_Carpeta_Descarga_Excels");

                if (p.ExisteParametroVigente("TOMATES_GUADIANA_Excel", DateTime.Now))
                {
                    ficheroLog.Add("Descargando el archivo: " + p.GetValue("TOMATES_GUADIANA_Excel"));
                    Console.WriteLine("Descargando el archivo: " + p.GetValue("TOMATES_GUADIANA_Excel"));

                    excel.Abrir(p.GetValue("TOMATES_GUADIANA_fileURL", DateTime.Now, DateTime.Now)
                        + p.GetValue("TOMATES_GUADIANA_Excel", DateTime.Now, DateTime.Now));

                    excel.Save(destino + p.GetValue("TOMATES_GUADIANA_Excel", DateTime.Now, DateTime.Now));

                    //sharePoint.fileURL = param.GetValue("AZUCARERA_fileURL", DateTime.Now, DateTime.Now)
                    //    + param.GetValue("AZUCARERA_Excel_BAÑEZA", DateTime.Now, DateTime.Now);
                    //sharePoint.destination = destino +
                    //    param.GetValue("AZUCARERA_Excel_BAÑEZA", DateTime.Now, DateTime.Now);
                    //sharePoint.Download();
                }



                ficheroLog.Add("Fin descarga de archivos Excel");
                ficheroLog.Add("================================");
                ficheroLog.Add("");
                hay_error = false;

            }
            catch (Exception e)
            {
                hay_error = true;
                ficheroLog.AddError("Automatismo_gas_Excel.DescargaArchivosExcel: " + e.Message);

                if (p.GetValue("utilizar_telegram") == "S")
                    telegram.SendMessage("CUIDADO ERROR!!!!"
                       + System.Environment.NewLine
                       + "================="
                       + System.Environment.NewLine
                       + "Necesario revisar proceso TOMATES DEL GUADIANA S.COOP. por ERROR en Automatismo_gas_Excel.DescargaArchivosExcel: "
                       + System.Environment.NewLine
                       + e.Message);
                
            }

        }

        private string GetLastActivationCode()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string code = "";

            try
            {
                strSql = "select productcode from atrgas_activaciones where"
                   + " cups = '" + p.GetValue("TOMATES_GUADIANA_CUPS") + "' and"
                   + " productenddate = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();                
                while (r.Read())
                {
                    code = r["productcode"].ToString();
                }
                db.CloseConnection();
                return code;
            }
            catch(Exception ex)
            {
                ficheroLog.AddError("Automatismo_gas_Excel.GetLastActivationCode: " + ex.Message);
                return "";
            }
        }
        private Int32 GetLastQd()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            Int32 qd = 0;

            try
            {
                strSql = "select productqd from atrgas_activaciones where"
                   + " cups = '" + p.GetValue("TOMATES_GUADIANA_CUPS") + "' and"
                   + " productenddate = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();                
                while (r.Read())
                {
                    qd = Convert.ToInt32(r["productqd"]);
                }
                db.CloseConnection();
                return qd;
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("Automatismo_gas_Excel.GetLastQd: " + ex.Message);
                return 0;
            }
        }

        private bool HayQueLanzarProceso(string proceso)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            DateTime hoy = new DateTime();
            bool lanzar = false;

            try
            {

                hoy = DateTime.Now;
                
                strSql = "select fecha from atrgas_fechas_procesos where"
                   + " proceso = '" + proceso + "'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                if (r.Read())
                {
                    lanzar = (Convert.ToDateTime(r["fecha"]).Date < hoy.Date);
                }

                db.CloseConnection();
                ficheroLog.Add("Se comprueba si hay que lanzar el proceso: "
                    + proceso + " con resultado = " + lanzar);
                return lanzar;
            }
            catch (Exception ex)
            {
                hay_error = true;
                ficheroLog.AddError("Automatismo_gas_Excel.HayQueLanzarProceso: " + ex.Message);
                return false;
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
            }
            catch (Exception e)
            {
                hay_error = true;
                if (p.GetValue("utilizar_telegram") == "S")
                    telegram.SendMessage("CUIDADO ERROR!!!!"
                    + System.Environment.NewLine
                    + "================="
                    + System.Environment.NewLine
                    + "Necesario revisar proceso Tomates del Guadiana por ERROR en BorrarContenidoDirectorio: "
                    + System.Environment.NewLine
                    + e.Message);
                ficheroLog.AddError("Azucarera.BorrarContenidoDirectorio: " + e.Message);
            }

        }

        private void ActualizaTablaMySQLProcesos(string proceso)
        {

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            DateTime ahora = new DateTime();

            // Actualizamos campos descripcion_autoconsumo de la tabla PS
            ahora = DateTime.Now;
            strSql = "update atrgas_fechas_procesos set fecha = '" + ahora.ToString("yyyy-MM-dd") + "'"
                + " where proceso = '" + proceso + "'";

            Console.WriteLine("Actualizamos la fecha de los procesos de la tabla en MySQL");
            ficheroLog.Add("Actualizamos la fecha de los procesos de la tabla en MySQL");
            Console.WriteLine(strSql);
            ficheroLog.Add(strSql);

            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

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

                if (lista.Count > 0)
                {


                    from = p.GetValue("remitente");
                    cc = p.GetValue("remitente");
                    subject = p.GetValue("TOMATES_GUADIANA_mail_asunto_distribuidora")
                        + " " + DateTime.Now.AddDays(1).ToString("dd/MM/yyyy");

                    attachment = System.Environment.CurrentDirectory + p.GetValue("mail_image_path", DateTime.Now, DateTime.Now);

                    //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                    EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                    // Aviso al gestor
                    body = GeneraCuerpoHTMLGestores(lista);
                    to = p.GetValue("TOMATES_GUADIANA_Mails_Aviso_Gestores");
                    subject = p.GetValue("TOMATES_GUADIANA_mail_asunto_copia_gestor") + DateTime.Now.AddDays(1).ToString("yyyyMMdd");

                    string[] listaArchivosExcel = Directory.GetFiles(p.GetValue("TOMATES_GUADIANA_Carpeta_Descarga_Excels", DateTime.Now, DateTime.Now), "*.xlsx");
                    for (int x = 0; x < listaArchivosExcel.Length; x++)
                    {
                        FileInfo file = new FileInfo(listaArchivosExcel[x]);
                        attachment += ";" + file.FullName;
                    }


                    //mes.SendMailWeb(from, to, cc, subject, body, attachment);
                    mes.SendMail(from, to, cc, subject, body, attachment);
                    hay_error = false;

                }

            }
            catch (Exception e)
            {
                hay_error = true;
                ficheroLog.AddError("Automatismo_Gas_Excels.GeneraCorreo: " + e.Message);
            }
        }
        public string GeneraCuerpoHTMLGestores(List<EndesaEntity.contratacion.gas.Solicitud> listaSol)
        {
            string body = "";
            string linea = "";

            try
            {

                body = p.GetValue("html_body_head_gestor", DateTime.Now, DateTime.Now);
                for (int i = 0; i < listaSol.Count; i++)
                {
                    linea = p.GetValue("html_body_line_gestor", DateTime.Now, DateTime.Now);
                    linea = linea.Replace("CUPS", listaSol[i].cups);
                    linea = linea.Replace("PLANTA", listaSol[i].descripcion);
                    linea = linea.Replace("CAUDAL_SOLICITADA", listaSol[i].detalle[0].qd.ToString("N0"));
                    linea = linea.Replace("FECHA_INICIO_SOLICITADA", listaSol[i].detalle[0].fecha_inicio.ToString("dd/MM/yyyy"));
                    linea = linea.Replace("FECHA_FIN_SOLICITADA", listaSol[i].detalle[0].fecha_fin.ToString("dd/MM/yyyy"));
                    linea = linea.Replace("PRODUCTO_SOLICITADO", listaSol[i].detalle[0].producto);
                    body += linea;
                }
                body += p.GetValue("html_body_foot_gestor", DateTime.Now, DateTime.Now);

                if (DateTime.Now.Hour > 14)
                    body = body.Replace("Buenos d&iacute;as.", "Buenas tardes:");
                return body;
            }
            catch (Exception e)
            {
                ficheroLog.AddError("GeneraCuerpoHTMLGestores: " + e.Message);
                hay_error = true;
                return "";
            }
        }
     
        private void EnvioMail_DatosFaltantesExcel()
        {
            string body = "";
            string subject = "";
            string from = "";
            string to = "";
            string cc = null;
            string attachment = null;

            from = p.GetValue("remitente");
            to = p.GetValue("TOMATES_GUADIANA_Mails_Aviso_Gestores_Error");
            cc = p.GetValue("remitente");

            body = "Se ha detectado qd sin informar para el día " + DateTime.Now.AddDays(1).ToString("dd/MM/yyyy")
                + System.Environment.NewLine
                + "Un saludo.";
            
            subject = "Qd sin informar para el día " + DateTime.Now.AddDays(1).ToString("dd/MM/yyyy");


            //EndesaBusiness.mail.MailExchangeServer mes =
            //    new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");

            EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
            mes.SendMail(from, to, cc, subject, body, attachment);

            if (p.GetValue("utilizar_telegram") == "S")
                telegram.SendMessage("Proceso Tomates del Guadiana:  " 
                    + subject);
        }

        public void Subir_a_FTP_Extremadura()
        {
            EndesaBusiness.utilidades.UltimateFTP ftp;
            string ruta_destino = "";
            string distribuidora = "";

            try
            {
                string[] listaArchivos = Directory.GetFiles(p.GetValue("TOMATES_GUADIANA_inbox", DateTime.Now, DateTime.Now), "*.xml");
                if (listaArchivos.Length > 0)
                {

                    ficheroLog.Add("Conectando al FTP: " + p.GetValue("ftp_extremadura_server", DateTime.Now, DateTime.Now));
                    ftp = new EndesaBusiness.utilidades.UltimateFTP(
                        p.GetValue("ftp_extremadura_server", DateTime.Now, DateTime.Now),
                        p.GetValue("ftp_extremadura_user", DateTime.Now, DateTime.Now),
                        utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_extremadura_pass", DateTime.Now, DateTime.Now), true),
                        p.GetValue("ftp_extremadura_port", DateTime.Now, DateTime.Now));


                    for (int i = 0; i < listaArchivos.Length; i++)
                    {
                        FileInfo fichero = new FileInfo(listaArchivos[i]);
                        ficheroLog.Add("Detectado: " + fichero.Name);
                        distribuidora = fichero.Name.Substring(16, 4);
                        ruta_destino = p.GetValue("ftp_extremadura_ruta_destino", DateTime.Now, DateTime.Now);
                        ftp.UploadInSecureFTP(p.GetValue("ftp_extremadura_server", DateTime.Now, DateTime.Now) + fichero.Name, listaArchivos[i]);

                        if (p.GetValue("utilizar_telegram") == "S")
                            telegram.SendMessage("Se publica en el FTP el archivo " + fichero.Name + " en la ruta " + ruta_destino);

                        ficheroLog.Add("Se publica en el FTP el archivo " + fichero.Name + " en la ruta " + ruta_destino);

                    }
                }
                error_ftp = false;
            }
            catch (Exception e)
            {
                error_ftp = true;
                ficheroLog.AddError("Subir_a_FTP_Extremadura: " + e.Message);

                if (p.GetValue("utilizar_telegram") == "S")
                    telegram.SendMessage("CUIDADO ERROR!!!!"
                       + System.Environment.NewLine
                       + "================="
                       + System.Environment.NewLine
                       + "Necesario revisar proceso por ERROR: "
                       + System.Environment.NewLine
                       + e.Message);



            }
        }


        private DateTime UltimoDiaAcepetacion()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            DateTime ultimo_dia = new DateTime();

            
            
            strSql = "select max(communicationsdate) max_communicationsdate from atrgas_aceptaciones";


            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                ultimo_dia = Convert.ToDateTime(r["max_communicationsdate"]);
            }
            db.CloseConnection();
            return ultimo_dia;
        }



    }
}
