using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using EndesaBusiness.servidores;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using Microsoft.Graph;

namespace EndesaBusiness.facturacion
{
    public class Informe_Contratos_Complementos
    {

        utilidades.Param param;
        logs.Log ficheroLog;
        EndesaBusiness.utilidades.Seguimiento_Procesos sp;


        Dictionary<string, EndesaEntity.facturacion.InformeContratosComplementos> dic_foto;
        Dictionary<string, List<EndesaEntity.contratacion.ComplementosContrato>> dic_complementos;
        EndesaBusiness.contratacion.PS_AT psat;


        public Informe_Contratos_Complementos()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Contratos_Complementos");
            param = new utilidades.Param("cont_comp_param", servidores.MySQLDB.Esquemas.CON);
            sp = new utilidades.Seguimiento_Procesos();
        }

        public void Proceso()
        {

            Console.WriteLine("INFORME CONTRATOS COMPLEMENTOS");
            Console.WriteLine("==============================");

            ficheroLog.Add("INFORME CONTRATOS COMPLEMENTOS");
            ficheroLog.Add("==============================");

            string archivo_informe = "";
            EndesaBusiness.utilidades.Fechas f = new utilidades.Fechas();           
            

            DateTime inicio = new DateTime();
            DateTime fin = new DateTime();
            DateTime ultima_ejecucion = new DateTime();

            try
            {
                ultima_ejecucion =
                sp.GetFecha_FinProceso("Facturación", "Informe Contratos Complementos", "4_Envio_Informe");

                Console.WriteLine("Última Ejecución del informe: " + ultima_ejecucion.ToString("dd/MM/yyyy HH:mm:ss"));

                if (ultima_ejecucion < DateTime.Now.Date)
                {
                    //psat = new contratacion.PS_AT(true);                    

                     DescargaFicheroExtractorCalendarios();

                    Console.WriteLine("Última ejecución del proceso: "
                        + ultima_ejecucion.ToString("dd/MM/yyyy")
                        + " < " + DateTime.Now.Date.ToString("dd/MM/yyyy"));

                    //string archivo = param.GetValue("ruta_salida")
                    //+ param.GetValue("COMP_CONT_PREF_ARCH1")
                    //+ f.UltimoDiaHabil().ToString("MMdd")
                    //+ param.GetValue("COMP_CONT_PREF_ARCH1_EXT");

                    string archivo = param.GetValue("ruta_salida")
                    + param.GetValue("COMP_CONT_PREF_ARCH2")                    
                    + param.GetValue("COMP_CONT_PREF_ARCH1_EXT");

                    FileInfo fileInfo = new FileInfo(archivo);
                    if (fileInfo.Exists)                    {

                        ImportarArchivoContratoComplementos(archivo);
                    

                        //dic_foto = CargaContratos(null);
                        dic_foto = CargaContratosComplementosCalendarios(null);
                        //dic_complementos = CargaContratosComplementosCalendarios_v2();

                        sp.Update_Fecha_Inicio("Facturación", "Informe Contratos Complementos", "3_Ejecución Informe");
                        archivo_informe = GeneraInforme();
                        sp.Update_Fecha_Fin("Facturación", "Informe Contratos Complementos", "3_Ejecución Informe");


                        sp.Update_Fecha_Inicio("Facturación", "Informe Contratos Complementos", "4_Envio_Informe");
                        if (param.GetValue("EnviarMail") == "S")
                            EnvioCorreo(archivo_informe);
                        sp.Update_Fecha_Fin("Facturación", "Informe Contratos Complementos", "4_Envio_Informe");

                        fileInfo.Delete();
                    }
                    else
                    {
                        EnvioCorreoNoExisteArchivo(archivo);
                    }

                }
                
            }
            catch(Exception e)
            {
                ficheroLog.AddError("Proceso: " + e.Message);
            }

            


        }

        public void DescargaFicheroExtractor()
        {
            EndesaBusiness.utilidades.Fechas f = new utilidades.Fechas();
            string md5 = "";
            string mensaje = "";

            try
            {

                string archivo_origen = param.GetValue("ruta_salida")
                    + param.GetValue("COMP_CONT_PREF_FICHERO_EXTRACTOR");


                string archivo_destino = param.GetValue("ruta_salida")
                    + param.GetValue("COMP_CONT_PREF_ARCH1")
                    + f.UltimoDiaHabil().ToString("MMdd")
                    + param.GetValue("COMP_CONT_PREF_ARCH1_EXT");

                FileInfo file = new FileInfo(archivo_origen);
                FileInfo file_destino = new FileInfo(archivo_destino);

                mensaje = "Llamando a extractor --> " +
                        param.GetValue("Extractor_Contratos_Complementos") + " para el dia " +
                        utilidades.Fichero.UltimoDiaHabil_YYMMDD();

                Console.WriteLine(mensaje);

                ficheroLog.Add(mensaje);

                sp.Update_Fecha_Inicio("Facturación", "Informe Contratos Complementos", "1_Ejecución Extractor");
                utilidades.Fichero.EjecutaComando(param.GetValue("Extractor_Contratos_Complementos"), null);
                

                Console.WriteLine("Fin extractor");
                ficheroLog.Add("Fin extractor");

                if (file_destino.Exists)
                    file_destino.Delete();


                if (file.Exists)
                {
                    string ultimo_md5 = param.GetValue("COMP_CONT_PREF_MD5_contratosPS");
                    md5 = utilidades.Fichero.checkMD5(file.FullName).ToString();

                    mensaje = "Ultimo MD5 procesado: "
                        + ultimo_md5
                        + " vs " + md5
                        + " --> " + (md5 != ultimo_md5);

                    if (file.Length > 0 && (md5 != ultimo_md5))
                    {
                        System.IO.File.Move(archivo_origen, archivo_destino);
                        param.code = "COMP_CONT_PREF_MD5_contratosPS";
                        param.from_date = new DateTime(2022, 08, 30);
                        param.to_date = new DateTime(4999, 12, 31);
                        param.value = md5;
                        param.Save();

                        sp.Update_Fecha_Fin("Facturación", "Informe Contratos Complementos", "1_Ejecución Extractor");

                    }
                    else
                        file.Delete();

                }


            }
            catch (Exception e)
            {
                ficheroLog.AddError("DescargaFicheroExtractor: " + e.Message);
            }

        }
        public void DescargaFicheroExtractorCalendarios()
        {
            EndesaBusiness.utilidades.Fechas f = new utilidades.Fechas();
            string md5 = "";
            string mensaje = "";

            try
            {

                

                string archivo_origen = param.GetValue("ruta_salida")
                    + param.GetValue("COMP_CONT_PREF_ARCH2")
                    + param.GetValue("COMP_CONT_PREF_ARCH1_EXT");


                FileInfo file = new FileInfo(archivo_origen);
                

                mensaje = "Llamando a extractor --> " +
                        param.GetValue("Extractor_Contratos_Complementos_Calendarios") + " para el dia " +
                        utilidades.Fichero.UltimoDiaHabil_YYMMDD();

                Console.WriteLine(mensaje);

                ficheroLog.Add(mensaje);

                sp.Update_Fecha_Inicio("Facturación", "Informe Contratos Complementos", "1_Ejecución Extractor");
                utilidades.Fichero.EjecutaComando(param.GetValue("Extractor_Contratos_Complementos_Calendarios"), null);

                Console.WriteLine("Fin extractor");
                ficheroLog.Add("Fin extractor");

                ficheroLog.Add("Buscando archivo: " + file.FullName);
                
                if (file.Exists)
                {
                    string ultimo_md5 = param.GetValue("MD5_archivo_contratosPS_Calendarios");
                    md5 = utilidades.Fichero.checkMD5(file.FullName).ToString();

                    mensaje = "Ultimo MD5 procesado: "
                        + ultimo_md5
                        + " vs " + md5
                        + " --> " + (md5 != ultimo_md5);

                    if (file.Length > 0 && (md5 != ultimo_md5))
                    {                        
                        param.code = "MD5_archivo_contratosPS_Calendarios";
                        param.from_date = new DateTime(2022, 08, 30);
                        param.to_date = new DateTime(4999, 12, 31);
                        param.value = md5;
                        param.Save();

     
                    }
                    
                }


            }
            catch (Exception e)
            {
                ficheroLog.AddError("DescargaFicheroExtractor: " + e.Message);
            }

        }

        private void EnvioCorreo(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            Dictionary<string, DateTime> dic = new Dictionary<string, DateTime>();

            try
            {

                dic.Add("PS_AT", sp.GetFecha_InicioProceso("Contratación", "PS_AT", "PS_AT"));
                dic.Add("Contratos Complementos Calendarios", param.LastUpdateParameter("MD5_archivo_contratosPS_Calendarios"));

                string from = param.GetValue("buzon_envio");
                string to = param.GetValue("email_para");
                string cc = param.GetValue("fac_inf_conv_cont_email_cc");
                string body = body = GeneraCuerpoHTML(dic);
                string subject = param.GetValue("email_asunto") + " a " + DateTime.Now.ToString("dd/MM/yyyy");
                                
                
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
                mes.SendMail(from, to, cc, subject, body,
                    System.Environment.CurrentDirectory + param.GetValue("mail_image_path"));

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("buzon_envio"));
                //fileInfo.Delete();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }
        private void EnvioCorreoNoExisteArchivo(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                string from = param.GetValue("buzon_envio");
                string to = param.GetValue("fac_inf_conv_cont_email_falta_archivo_para");
                //string cc = param.GetValue("fac_inf_conv_cont_email_falta_archivo_cc");
                string subject = param.GetValue("fac_inf_conv_cont_mail_falta_asunto");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  No se ha encontrado el archivo ").Append(fileInfo.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Sin este archivo de extracción no se puede generar el informe. ");

                textBody.Append("  Dentro de 15 minutos se volverá a comprobar su existencia. ");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");

                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                // EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
                EndesaBusiness.office.SendMail mes = new office.SendMail(from);

                List<string> lista_para = to.Split(';').ToList();
                //List<string> lista_cc = cc.Split(';').ToList();
                List<string> lista_adjuntos = archivo.Split(';').ToList();

                mes.para = lista_para;
                mes.cc = null;
                mes.asunto = subject;
                mes.htmlCuerpo = textBody.ToString();
                mes.adjuntos = lista_adjuntos;


                if (param.GetValue("EnviarMail") == "S")
                    // mes.SendMail(from, to, null, subject, textBody.ToString(), null);
                    mes.Send();

                else
                    // mes.SaveMail(from, to, null, subject, textBody.ToString(), null);
                    mes.Save();

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("buzon_envio"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }

        public string GeneraCuerpoHTML(Dictionary<string, DateTime> dic_extracciones)
        {
            string body = "";
            string linea = "";

            try
            {
                

                body = param.GetValue("html_head", DateTime.Now, DateTime.Now);
                foreach(KeyValuePair<string, DateTime> kvp in dic_extracciones)
                {
                    linea = param.GetValue("html_body", DateTime.Now, DateTime.Now);
                    linea = linea.Replace("extracion", kvp.Key);
                    linea = linea.Replace("fecha_extraccion", kvp.Value.ToString("dd/MM/yyyy"));
                    body += linea;
                }
                body += param.GetValue("html_foot", DateTime.Now, DateTime.Now);

                if (DateTime.Now.Hour > 14)
                    body = body.Replace("Buenos d&iacute;as:", "Buenas tardes:");
                return body;
            }
            catch (Exception e)
            {

                return "";
            }
        }


        private Dictionary<string, EndesaEntity.facturacion.InformeContratosComplementos> CargaContratos(string ccounips)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string tarifa = "";
            int faltacon = 0;
            int fpsercon = 0;
            int fprevbaj = 0;
            int fbajacon = 0;

            Dictionary<string, EndesaEntity.facturacion.InformeContratosComplementos> d
                = new Dictionary<string, EndesaEntity.facturacion.InformeContratosComplementos>();

            db = new MySQLDB(MySQLDB.Esquemas.CON);
            try
            {



                strSql = "SELECT ps.IDU AS CCOUNIPS, ps.CUPS22, ps.CONTREXT, ps.estadoCont AS TESTCONT,"
                    + " ps.fAltaCont AS FALTACON, ps.FPSERCON, ps.fPrevBajaCont AS FPREVBAJ,"
                    + " ps.fBajaCont AS FBAJACON, cc.CEMPTITU, cc.CCONTRPS, cc.CNUMSCPS, cc.CLINNEG,"
                    + " cc.CCLIENTE, cc.CCALENPO, cc.VDIAFACT, cc.FSIGFACT,"
                    + " cc.FFINVESU, cc.CTARIFA, cc.CUPSREE, cc.TPUNTMED, cc.CCOMPOBJ, cc.VNSEGHOR, cc.VNUMTRAM," 
                    + " cc.VPARAM01, cc.VPARAM02, cc.VPARAM03, cc.VPARAM04, cc.VPARAM05, cc.CCOMAUTO, cc.TTICONPS," 
                    + " cc.TPOTENCIP1, cc.VPOTCALIP1, cc.TPOTENCIP2, cc.VPOTCALIP2, cc.TPOTENCIP3, cc.VPOTCALIP3," 
                    + " cc.TPOTENCIP4, cc.VPOTCALIP4, cc.TPOTENCIP5,cc.VPOTCALIP5, cc.TPOTENCIP6, cc.VPOTCALIP6,"
                    + " ps.TENSION"
                    + " FROM PS_AT ps"
                    + " LEFT OUTER JOIN cont_comp_contratos cc ON"
                    + " cc.CCOUNIPS = ps.IDU AND"
                    + " cc.CONTREXT = ps.CONTREXT";

                if (ccounips != null)
                    strSql += " where ps.IDU = '" + ccounips + "'";


                
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    EndesaEntity.facturacion.InformeContratosComplementos o;
                    if (!d.TryGetValue(r["CCOUNIPS"].ToString(), out o))
                    {

                        tarifa = r["CTARIFA"].ToString();

                        if (r["FALTACON"] != System.DBNull.Value)
                            faltacon = Convert.ToInt32(r["FALTACON"]);

                        if (r["FPSERCON"] != System.DBNull.Value)
                            fpsercon = Convert.ToInt32(r["FPSERCON"]);

                        if (r["FPREVBAJ"] != System.DBNull.Value)
                            fprevbaj = Convert.ToInt32(r["FPREVBAJ"]);

                        if (r["FBAJACON"] != System.DBNull.Value)
                            fbajacon = Convert.ToInt32(r["FBAJACON"]);

                        o = new EndesaEntity.facturacion.InformeContratosComplementos();
                        o.cups = r["CCOUNIPS"].ToString();

                        if (r["CUPS22"] != System.DBNull.Value)
                        {
                            if (r["CUPS22"].ToString().Length == 22)
                            {
                                o.cups20 = r["CUPS22"].ToString().Substring(0, 20);
                                o.cups22 = r["CUPS22"].ToString();
                            }                            
                        }


                        if (r["CONTREXT"] != System.DBNull.Value)
                            o.contrato = r["CONTREXT"].ToString();

                        if (r["CNUMSCPS"] != System.DBNull.Value)
                            o.version = Convert.ToInt32(r["CNUMSCPS"]);

                        if (r["TENSION"] != System.DBNull.Value)
                            o.tension = Convert.ToInt32(r["TENSION"]);

                        o.dic = Cabecera(tarifa, faltacon, fpsercon, fprevbaj, fbajacon);

                        if(r["CCOMPOBJ"] == System.DBNull.Value)                       
                            o.dic.Add("Sin Datos Contrato", "Sin Datos Contrato");                        


                        //d.Add(o.contrato, o);
                        d.Add(o.cups, o);

                    }
                    else
                    {
                        string oo;
                        if (r["CCOMPOBJ"] != System.DBNull.Value)
                        {
                            oo = r["CCOMPOBJ"].ToString();
                            o.dic.Add(oo, oo);
                        }
                        
                    }


                }
                r.Close();
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                db.CloseConnection();
                ficheroLog.AddError("CargaContratos: " + e.Message);
                return null;
            }
        }

        private Dictionary<string, EndesaEntity.facturacion.InformeContratosComplementos> CargaContratosComplementosCalendarios(string ccounips)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string tarifa = "";
            int faltacon = 0;
            int fpsercon = 0;
            int fprevbaj = 0;
            int fbajacon = 0;

            Dictionary<string, EndesaEntity.facturacion.InformeContratosComplementos> d
                = new Dictionary<string, EndesaEntity.facturacion.InformeContratosComplementos>();

            db = new MySQLDB(MySQLDB.Esquemas.CON);
            try
            {



                strSql = "SELECT ps.IDU AS CCOUNIPS, ps.CUPS22, ps.CONTREXT, ps.estadoCont AS TESTCONT,"
                    + " ps.fAltaCont AS FALTACON, ps.FPSERCON, ps.fPrevBajaCont AS FPREVBAJ,"
                    + " ps.fBajaCont AS FBAJACON, cc.CEMPTITU, cc.CCONTRPS, cc.CNUMSCPS, cc.CLINNEG,"
                    + " cc.CCLIENTE, cc.CCALENPO, "
                    + " cc.CTARIFA,  cc.CCOMPOBJ, cc.VNSEGHOR,"
                    + " cc.VPARAM01, cc.VPARAM02, cc.VPARAM03, cc.VPARAM04, cc.VPARAM05,"
                    + " ps.TENSION"
                    + " FROM PS_AT ps"
                    + " LEFT OUTER JOIN cont_comp_contratos_calendarios cc ON"
                    + " cc.CCOUNIPS = ps.IDU";
                    //+ " cc.CONTREXT = ps.CONTREXT";

                if (ccounips != null)
                    strSql += " where ps.IDU = '" + ccounips + "'";

                strSql += " group by ps.IDU, cc.CCOMPOBJ";

                
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    EndesaEntity.facturacion.InformeContratosComplementos o;
                    if (!d.TryGetValue(r["CCOUNIPS"].ToString(), out o))
                    {

                        tarifa = r["CTARIFA"].ToString();

                        if (r["FALTACON"] != System.DBNull.Value)
                            faltacon = Convert.ToInt32(r["FALTACON"]);

                        if (r["FPSERCON"] != System.DBNull.Value)
                            fpsercon = Convert.ToInt32(r["FPSERCON"]);

                        if (r["FPREVBAJ"] != System.DBNull.Value)
                            fprevbaj = Convert.ToInt32(r["FPREVBAJ"]);

                        if (r["FBAJACON"] != System.DBNull.Value)
                            fbajacon = Convert.ToInt32(r["FBAJACON"]);

                        o = new EndesaEntity.facturacion.InformeContratosComplementos();
                        o.cups = r["CCOUNIPS"].ToString();

                        if (r["CUPS22"] != System.DBNull.Value)
                        {
                            if (r["CUPS22"].ToString().Length == 22)
                            {
                                o.cups20 = r["CUPS22"].ToString().Substring(0, 20);
                                o.cups22 = r["CUPS22"].ToString();
                            }
                        }


                        if (r["CONTREXT"] != System.DBNull.Value)
                            o.contrato = r["CONTREXT"].ToString();

                        if (r["CNUMSCPS"] != System.DBNull.Value)
                            o.version = Convert.ToInt32(r["CNUMSCPS"]);

                        if (r["TENSION"] != System.DBNull.Value)
                            o.tension = Convert.ToInt32(r["TENSION"]);

                        o.dic = Cabecera(tarifa, faltacon, fpsercon, fprevbaj, fbajacon);

                        if (r["CCOMPOBJ"] == System.DBNull.Value)
                            o.dic.Add("Sin Datos Contrato", "Sin Datos Contrato");
                        else
                        {
                            string oo;
                            oo = r["CCOMPOBJ"].ToString();
                            o.dic.Add(oo, oo);
                            
                        }
                            


                        //d.Add(o.contrato, o);                        
                        d.Add(o.cups, o);

                    }
                    else
                    {
                        string oo;
                        if (r["CCOMPOBJ"] != System.DBNull.Value)
                        {
                            oo = r["CCOMPOBJ"].ToString();
                            o.dic.Add(oo, oo);
                        }

                    }


                }
                r.Close();
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                db.CloseConnection();
                ficheroLog.AddError("CargaContratos: " + e.Message);
                return null;
            }
        }
        private Dictionary<string, List<EndesaEntity.contratacion.ComplementosContrato>> CargaContratosComplementosCalendarios_v2()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";            

            Dictionary<string, List<EndesaEntity.contratacion.ComplementosContrato>> d =
            new Dictionary<string, List<EndesaEntity.contratacion.ComplementosContrato>>();
                
                

            db = new MySQLDB(MySQLDB.Esquemas.CON);
            try
            {

                strSql = "SELECT cc.CCOUNIPS,"                    
                    + " cc.CEMPTITU, cc.CCONTRPS, cc.CNUMSCPS, cc.CLINNEG,"
                    + " cc.CCLIENTE, cc.CCALENPO, "
                    + " cc.CTARIFA,  cc.CCOMPOBJ, cc.VNSEGHOR,"
                    + " cc.VPARAM01, cc.VPARAM02, cc.VPARAM03, cc.VPARAM04, cc.VPARAM05,"                    
                    + " FROM cont_comp_contratos_calendarios cc";

                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    EndesaEntity.contratacion.ComplementosContrato c =
                        new EndesaEntity.contratacion.ComplementosContrato();

                    c.cups13 = r["CCOUNIPS"].ToString();
                    c.codigo_complemento = r["CCOMPOBJ"].ToString();

                    for (int j = 0; j <= 4; j++)
                        c.vparam[j] = Convert.ToDouble(r["VPARAM0" + (j + 1)]);
                    

                    List<EndesaEntity.contratacion.ComplementosContrato> o;
                    if (!d.TryGetValue(r["CCOUNIPS"].ToString(), out o))
                    {
                        o = new List<EndesaEntity.contratacion.ComplementosContrato>();
                        o.Add(c);
                        d.Add(c.cups13, o);

                    }
                    else
                        o.Add(c);
                        

                }
                r.Close();
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                db.CloseConnection();
                ficheroLog.AddError("CargaContratos: " + e.Message);
                return null;
            }
        }

        private Dictionary<string, string>
            Cabecera(string tarifa, int faltacon, int fpsercon, int fprevbaj, int fbajacon)
        {

            Dictionary<string, string> d
                = new Dictionary<string, string>();

                       
            
            d.Add("TARIFA", tarifa);
            
            d.Add("FALTACON", faltacon > 0 ? faltacon.ToString() : "");
            
            d.Add("FPSERCON", fpsercon > 0 ? fpsercon.ToString() : "");
           
            d.Add("FPREVBAJ", fprevbaj > 0 ? fprevbaj.ToString() : "");
            
            d.Add("FBAJACON", fbajacon > 0 ? fbajacon.ToString() : "");

            return d;

        }

        private string GeneraInforme()
        {

            ExcelPackage excelPackage;
            FileInfo fileInfo;
            string variable = "";

            int f = 1;
            int c = 1;

            try
            {

                fileInfo = new FileInfo(param.GetValue("ruta_salida_sharepoint")
                + param.GetValue("nombre_informe")
                + DateTime.Now.ToString("yyyy_MM_dd_HHmmss") + ".xlsx");

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("Complementos");

                var headerCells = workSheet.Cells[1, 1, 1, 8];
                var headerFont = headerCells.Style.Font;

                var allCells = workSheet.Cells[1, 1, 1, 8];
                var cellFont = allCells.Style.Font;
                cellFont.Bold = true;

                workSheet.Cells[f, c].Value = "CUPS";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "CUPS20";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "CUPS22";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "CONTRATO";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;                

                workSheet.Cells[f, c].Value = "VERSION";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "VARIABLE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;
                
                workSheet.Cells[f, c].Value = "VALOR";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;
                                             

                foreach (KeyValuePair<string, EndesaEntity.facturacion.InformeContratosComplementos> p in dic_foto)
                {

                    foreach (KeyValuePair<string, string> pp in p.Value.dic)
                    {
                        c = 1;
                        f++;

                        workSheet.Cells[f, c].Value = p.Value.cups; c++;
                        workSheet.Cells[f, c].Value = p.Value.cups20; c++;
                        workSheet.Cells[f, c].Value = p.Value.cups22; c++;
                        workSheet.Cells[f, c].Value = p.Value.contrato; c++;
                        workSheet.Cells[f, c].Value = p.Value.version; c++;                        

                        switch (pp.Key)
                        {
                            case "TARIFA":
                                variable = pp.Key;
                                break;
                            case "FALTACON":
                                variable = pp.Key;
                                break;
                            case "FPSERCON":
                                variable = pp.Key;
                                break;
                            case "FPREVBAJ":
                                variable = pp.Key;
                                break;
                            case "FBAJACON":
                                variable = pp.Key;
                                break;
                            case "Sin Datos Contrato":
                                variable = "Sin Datos Contrato";
                                break;
                            default:
                                variable = "Complemento";
                                break;

                        }
                        workSheet.Cells[f, c].Value = variable; c++;
                        workSheet.Cells[f, c].Value = pp.Value; c++;                        

                    }

                }


                allCells = workSheet.Cells[1, 1, f, 8];

                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:G1"].AutoFilter = true;
                allCells.AutoFitColumns();

                excelPackage.Save();

                return fileInfo.FullName;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        private void ImportarArchivoContratoComplementos(string archivo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            string line = "";
            string[] campos;
            int i = 0;
            int c = 0;
            string strSql = "";
            int total_registros = 0;

            try
            {
                sp.Update_Fecha_Inicio("Facturación", "Informe Contratos Complementos", "2_Importación");

                strSql = "delete from cont_comp_contratos_calendarios";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                ficheroLog.Add("Importando archivo " + archivo);

                System.IO.StreamReader file = new System.IO.StreamReader(archivo, System.Text.Encoding.GetEncoding(1252));
                while ((line = file.ReadLine()) != null)
                {
                    campos = line.Split(';');
                    i++;
                    c = 1;
                    total_registros++;

                    if (firstOnly)
                    {
                        sb.Append("replace into cont_comp_contratos_calendarios");
                        sb.Append(" (CCOUNIPS, CONTREXT, TESTCONT, FALTACON,");
                        sb.Append(" CEMPTITU, CCONTRPS, CNUMSCPS, CLINNEG, CCLIENTE, CCALENPO,");
                        sb.Append(" CSEGMERC, CTARIFA, CCOMPOBJ, VNSEGHOR, VPARAM01,");
                        sb.Append(" VPARAM02, VPARAM03, VPARAM04, VPARAM05) values ");                        
                        
                        firstOnly = false;
                    }

                    sb.Append("('").Append(campos[c]).Append("',"); c++; // CCOUNIPS
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // CONTREXT

                    for (int j = 1; j <= 3; j++)
                    {
                        sb.Append(CN(campos[c])).Append(","); c++;
                    }

                    sb.Append(CN(campos[c])).Append(","); c++; // CCONTRPS

                    for (int j = 1; j <= 4; j++)
                    {
                        sb.Append(CN(campos[c])).Append(","); c++;
                    }

                    sb.Append("'").Append(campos[c]).Append("',"); c++; // CSEGMERC
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // CTARIFA                                   
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // CCOMPOBJ
                    sb.Append(CN(campos[c])); c++; // VNSEGHOR
                                        
                    
                    sb.Append(",").Append(CDouble(campos[c])); c++; // VPARAM01
                    sb.Append(",").Append(CDouble(campos[c])); c++; // VPARAM02
                    sb.Append(",").Append(CDouble(campos[c])); c++; // VPARAM03
                    sb.Append(",").Append(CDouble(campos[c])); c++; // VPARAM04
                    sb.Append(",").Append(CDouble(campos[c])); c++; // VPARAM05
                    sb.Append("),");


                    if (i == 250)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        i = 0;
                    }

                }
                file.Close();

                if (i > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                }

                ficheroLog.Add("Fin importación");
                ficheroLog.Add("Se han importado " + total_registros + " registros");
                FileInfo archivo_info = new FileInfo(archivo);
                archivo_info.Delete();

                sp.Update_Fecha_Fin("Facturación", "Informe Contratos Complementos", "2_Importación");

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void ImportarArchivoContratoComplementosCalendarios(string archivo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            string line = "";
            string[] campos;
            int i = 0;
            int c = 0;
            string strSql = "";
            int total_registros = 0;

            try
            {
                sp.Update_Fecha_Inicio("Facturación", "Informe Contratos Complementos", "3_Import_contratos_complementos_calendarios");

                strSql = "delete from cont_comp_contratos_calendarios";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                ficheroLog.Add("Importando archivo " + archivo);

                System.IO.StreamReader file = new System.IO.StreamReader(archivo, System.Text.Encoding.GetEncoding(1252));
                while ((line = file.ReadLine()) != null)
                {
                    campos = line.Split(';');
                    i++;
                    c = 1;
                    total_registros++;

                    if (firstOnly)
                    {
                        sb.Append("replace into cont_comp_contratos_calendarios");
                        sb.Append(" (CCOUNIPS, CONTREXT, TESTCONT, FALTACON,");
                        sb.Append(" CEMPTITU, CCONTRPS, CNUMSCPS, CLINNEG, CCLIENTE, CCALENPO,");
                        sb.Append(" CSEGMERC, CTARIFA, CCOMPOBJ, VNSEGHOR, VPARAM01,");
                        sb.Append(" VPARAM02, VPARAM03, VPARAM04, VPARAM05");                        
                        sb.Append(" ) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(campos[c]).Append("',"); c++; // CCOUNIPS
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // CONTREXT                  
                    sb.Append(CN(campos[c])).Append(","); c++; // TESTCONT
                    sb.Append(CF(campos[c])).Append(","); c++; // FALTACON
                    sb.Append(CN(campos[c])).Append(","); c++; // CEMPTITU

                    sb.Append(CN(campos[c])).Append(","); c++; // CCONTRPS

                    sb.Append(CN(campos[c])).Append(","); c++; // CNUMSCPS
                    sb.Append(CN(campos[c])).Append(","); c++; // CLINNEG
                    sb.Append(CN(campos[c])).Append(","); c++; // CCLIENTE
                    sb.Append(CN(campos[c])).Append(","); c++; // CCALENPO

                    sb.Append("'").Append(campos[c]).Append("',"); c++; // CSEGMERC
                    sb.Append(CN(campos[c])).Append(","); c++; // CTARIFA                                        
                    sb.Append(CN(campos[c])).Append(","); c++; // CCOMPOBJ
                    sb.Append(CF(campos[c])); c++; // VNSEGHOR
                    

                    for (int j = 1; j <= 5; j++)
                    {
                        sb.Append(",").Append(CDouble(campos[c])); c++;
                    }                   

                    sb.Append("),");


                    if (i == 250)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        i = 0;
                    }

                }
                file.Close();

                if (i > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                }

                ficheroLog.Add("Fin importación");
                ficheroLog.Add("Se han importado " + total_registros + " registros");
                FileInfo archivo_info = new FileInfo(archivo);
                archivo_info.Delete();

                sp.Update_Fecha_Fin("Facturación", "Informe Contratos Complementos", "3_Import_contratos_complementos_calendarios");

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        private string CF(string t)
        {
            if (t.Trim() == "00000000" || t.Trim() == "")
                return "null";
            else
                return t.Trim();
        }

        private string CN(string t)
        {
            if (t.Trim() == "00000000" || t.Trim() == "")
                return "'null'";
            else
                return "'" + t.Trim() + "'";
        }

        public string CDouble(String t)
        {
            t = t.Trim();
            if (t == "")
            {
                return "null";
            }
            else
            {
                t = t.Replace(" ", "");
                t = t.Replace("+", string.Empty);
                t = t.Replace("----------", string.Empty);
                t = t.Replace(".", string.Empty);
                t = t.Replace(",", ".");

                if (t == "")
                {
                    t = "null";
                }
            }

            return t;
        }
    }
}
