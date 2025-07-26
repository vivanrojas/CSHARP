using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion
{
    public class Covid
    {

        logs.Log ficheroLog;
        utilidades.Param p;

        public Covid()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_contratacion_covid");
            p = new utilidades.Param("covid_param", servidores.MySQLDB.Esquemas.CON);
                            
        }


        public void InformeBajasCovid()
        {
            int f = 1;
            int c = 1;

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            string fichero = "";
            bool hayDatos = false;

            try
            {

                fichero = p.GetValue("outbox_folder", DateTime.Now, DateTime.Now)
                    + p.GetValue("filename_sol_bajas", DateTime.Now, DateTime.Now)
                    + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";

                FileInfo fileInfo = new FileInfo(fichero);


                if (fileInfo.Exists)
                    fileInfo.Delete();

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("Solicitudes COVID");

                var headerCells = workSheet.Cells[1, 1, 1, 28];
                var headerFont = headerCells.Style.Font;
                f = 1;

                headerFont.Bold = true;

                #region Cabecera_Excel

                workSheet.Cells[f, c].Value = "CUPS20";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CUPS13";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CIF";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "Cliente";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CCONTRPS";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "Versión Contrato";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "Tarifa";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;


                workSheet.Cells[f, c].Value = "Estado Solicitud";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "Fecha Aceptación";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "Fecha Activación";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "Motivo Baja";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "Estado Contrato";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "Modelo fecha efecto";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "Fecha Solicitada";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                #endregion


                strSql = "SELECT substr(s.CUPS_EXT, 1,20) AS cups20,"
                    + " s.CCOUNIPS AS cups13, s.CIF AS cif, s.cliente," 
	                + " s.CCONTRPS AS ccontrps, s.VER_CONTR_PS AS version_contr_ps,"
	                + " s.TARIFA AS tarifa, s.estadoSolATR AS estadosolatr, s.fAcepRech,"
	                + " s.fActivacion, s.MOT_BAJA AS mot_baja, c.Descripcion AS EstadoContrato,"
	                + " s.MOD_FECHA_SOLCT, s.FECHA_SOLCT"
                    + " FROM cont.SOLATRMT_MAX s"
                    + " INNER JOIN cont.cont_estadoscontrato c ON"
                    + " c.Cod_Estado = lpad(s.TESTCONT, 3, '0')"
                    + " WHERE s.MOT_BAJA = 'CV'"
                    + " GROUP BY s.CUPS_EXT";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    hayDatos = true;

                    f++;
                    c = 1;

                    if (r["cups20"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["cups20"].ToString();
                    c++;
                    if (r["cups13"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["cups13"].ToString();
                    c++;
                    if (r["cif"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["cif"].ToString();
                    c++;
                    if (r["cliente"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["cliente"].ToString();
                    c++;
                    if (r["ccontrps"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["ccontrps"].ToString();
                    c++;
                    if (r["version_contr_ps"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["version_contr_ps"].ToString();
                    c++;
                    if (r["tarifa"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["tarifa"].ToString();
                    c++;
                    if (r["estadosolatr"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["estadosolatr"].ToString();
                    c++;
                    if (r["fAcepRech"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["fAcepRech"].ToString();
                    c++;
                    if (r["fActivacion"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["fActivacion"].ToString();
                    c++;
                    if (r["mot_baja"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["mot_baja"].ToString();
                    c++;
                    if (r["EstadoContrato"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["EstadoContrato"].ToString();
                    c++;
                    if (r["MOD_FECHA_SOLCT"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["MOD_FECHA_SOLCT"].ToString();
                    c++;
                    if (r["FECHA_SOLCT"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["FECHA_SOLCT"].ToString();
                    c++;



                }
                db.CloseConnection();

                var allCells = workSheet.Cells[1, 1, f, 28];
                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:J1"].AutoFilter = true;
                allCells.AutoFitColumns();

                excelPackage.Save();

                if (hayDatos)
                {
                    #region EnvioMail
                    office.SendMail s = new office.SendMail("Contratacionee@enel.com");
                    StringBuilder textBody = new StringBuilder();
                    s.para = utilidades.GOMail.Destinatarios("covid_fact", utilidades.GOMail.destinatarios.to);
                    s.cc = utilidades.GOMail.Destinatarios("covid_fact", utilidades.GOMail.destinatarios.cc);
                    s.asunto = "Listado solicitudes de Baja por situación COVID a " + DateTime.Now.ToString("dd/MM/yyyy");
                    s.adjunto = fileInfo.FullName;

                    textBody.Append(System.Environment.NewLine);
                    textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("   Adjuntamos el listado de contratos que han solicitado baja anticipada por situación crisis COVID.");
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("Un saludo.");

                    s.htmlCuerpo = textBody.ToString();
                    if (p.GetValue("EnviarMail", DateTime.Now, DateTime.Now) == "S")
                        s.Send();
                    else
                        s.Save();
                    #endregion
                }
                else
                {
                    #region EnvioMail
                    office.SendMail s = new office.SendMail("Contratacionee@enel.com");
                    StringBuilder textBody = new StringBuilder();
                    s.para = utilidades.GOMail.Destinatarios("covid_fact", utilidades.GOMail.destinatarios.to);
                    s.cc = utilidades.GOMail.Destinatarios("covid_fact", utilidades.GOMail.destinatarios.cc);
                    s.asunto = "Listado solicitudes de Baja por situación COVID a " + DateTime.Now.ToString("dd/MM/yyyy");                    

                    textBody.Append(System.Environment.NewLine);
                    textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("   No se ha detectado contratos que hayan solicitado baja anticipada por situación crisis COVID.");
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("Un saludo.");

                    s.htmlCuerpo = textBody.ToString();
                    if (p.GetValue("EnviarMail", DateTime.Now, DateTime.Now) == "S")
                        s.Send();
                    else
                        s.Save();
                    #endregion
                }


            }
            catch(Exception e)
            {
                
                ficheroLog.AddError("Covid.InformeBajasCovid --> " + e.Message);
                EndesaBusiness.mail.MailExchangeServer mes =
                    new EndesaBusiness.mail.MailExchangeServer(p.GetValue("usuario_errores", DateTime.Now, DateTime.Now));
                mes.SendMail(p.GetValue("errores_from", DateTime.Now, DateTime.Now),
                    p.GetValue("errores_to", DateTime.Now, DateTime.Now), null, "Error en Covid.InformeBajasCovid " + DateTime.Now.ToString("dd/MM/yyyy"),
                    System.Environment.NewLine +
                    System.Environment.NewLine +
                    e.Message, null);
            }

        }

        public void AplazadosQuePidenBaja()
        {
            int f = 1;
            int c = 1;

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            string fichero = "";
            bool hayDatos = false;

            try
            {

                fichero = p.GetValue("outbox_folder", DateTime.Now, DateTime.Now)
                    + p.GetValue("filename_sol_bajas_apla", DateTime.Now, DateTime.Now)
                    + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";

                FileInfo fileInfo = new FileInfo(fichero);


                if (fileInfo.Exists)
                    fileInfo.Delete();

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("Sol COVID BAJAS");

                var headerCells = workSheet.Cells[1, 1, 1, 28];
                var headerFont = headerCells.Style.Font;
                f = 1;

                headerFont.Bold = true;

                #region Cabecera_Excel

                workSheet.Cells[f, c].Value = "CUPS20";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "Empresa";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "NIF";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "Cliente";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "TARIFA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "DDISTRIB";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "Estado Contrato";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;


                workSheet.Cells[f, c].Value = "F. Alta Contrato";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "FPSERCON";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;
                

                #endregion



                strSql = "SELECT fp.*"
                    + " FROM cont.covid_fraccionamientopago fp"
                    + " INNER JOIN cont.SOLATRMT_MAX m ON"
                    + " SUBSTR(m.CUPS_EXT, 1, 20) = fp.CUPS20"
                    + " WHERE m.tipoSolATR = 'BAJA'";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    hayDatos = true;

                    f++;
                    c = 1;

                    if (r["CUPS20"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CUPS20"].ToString();
                    c++;
                    if (r["Empresa"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["Empresa"].ToString();
                    c++;
                    if (r["NIF"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["NIF"].ToString();
                    c++;
                    if (r["Cliente"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["Cliente"].ToString();
                    c++;
                    if (r["TARIFA"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                    c++;
                    if (r["DDISTRIB"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["DDISTRIB"].ToString();
                    c++;
                    if (r["estadoContrato"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["estadoContrato"].ToString();
                    c++;
                    if (r["fAltaCont"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["fAltaCont"].ToString();
                    c++;
                    if (r["FPSERCON"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["FPSERCON"].ToString();


                }
                db.CloseConnection();

                var allCells = workSheet.Cells[1, 1, f, 28];
                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:J1"].AutoFilter = true;
                allCells.AutoFitColumns();

                excelPackage.Save();

                if (hayDatos)
                {
                    #region EnvioMail
                    office.SendMail s = new office.SendMail("Contratacionee@enel.com");
                    StringBuilder textBody = new StringBuilder();
                    s.para = utilidades.GOMail.Destinatarios("covid_solBaj", utilidades.GOMail.destinatarios.to);
                    s.cc = utilidades.GOMail.Destinatarios("covid_solBaj", utilidades.GOMail.destinatarios.cc);
                    s.asunto = "Listado solicitudes de Baja de contrato por situación COVID a " + DateTime.Now.ToString("dd/MM/yyyy");
                    s.adjunto = fileInfo.FullName;

                    textBody.Append(System.Environment.NewLine);
                    textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("   Adjuntamos el listado de contratos que han solicitado baja y aplazamiento de la deuda.");
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("Un saludo.");

                    s.htmlCuerpo = textBody.ToString();
                    if (p.GetValue("EnviarMail", DateTime.Now, DateTime.Now) == "S")
                        s.Send();
                    else
                        s.Save();
                    #endregion
                }
                else
                {
                    #region EnvioMail
                    office.SendMail s = new office.SendMail("Contratacionee@enel.com");
                    StringBuilder textBody = new StringBuilder();
                    s.para = utilidades.GOMail.Destinatarios("covid_solBaj", utilidades.GOMail.destinatarios.to);
                    s.cc = utilidades.GOMail.Destinatarios("covid_solBaj", utilidades.GOMail.destinatarios.cc);
                    s.asunto = "Listado solicitudes de Baja de contrato por situación COVID a " + DateTime.Now.ToString("dd/MM/yyyy");                    

                    textBody.Append(System.Environment.NewLine);
                    textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("   No se han detectado contratos que hayan solicitado baja y aplazamiento de la deuda.");
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("Un saludo.");

                    s.htmlCuerpo = textBody.ToString();
                    if (p.GetValue("EnviarMail", DateTime.Now, DateTime.Now) == "S")
                        s.Send();
                    else
                        s.Save();
                    #endregion
                }



            }
            catch(Exception e)
            {

                ficheroLog.AddError("Covid.AplazadosQuePidenBaja --> " + e.Message);
                EndesaBusiness.mail.MailExchangeServer mes = 
                    new EndesaBusiness.mail.MailExchangeServer(p.GetValue("usuario_errores", DateTime.Now, DateTime.Now));
                mes.SendMail(p.GetValue("errores_from", DateTime.Now, DateTime.Now),
                    p.GetValue("errores_to", DateTime.Now, DateTime.Now), null, "Error en Covid.AplazadosQuePidenBaja " + DateTime.Now.ToString("dd/MM/yyyy"),
                    System.Environment.NewLine +
                    System.Environment.NewLine +
                    e.Message, null);
            }
        }

        public void SolatrMT_Covid19()
        {
            int f = 1;
            int c = 1;

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            string fichero = "";
            bool hayDatos = false;

            DateTime fecha = new DateTime();

            try
            {

                fichero = p.GetValue("outbox_folder", DateTime.Now, DateTime.Now)
                    + p.GetValue("filename_solatrmt", DateTime.Now, DateTime.Now)
                    + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";

                FileInfo fileInfo = new FileInfo(fichero);


                if (fileInfo.Exists)
                    fileInfo.Delete();

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("Soliticitudes");

                var headerCells = workSheet.Cells[1, 1, 1, 28];
                var headerFont = headerCells.Style.Font;
                f = 1;

                headerFont.Bold = true;

                #region Cabecera_Excel

                workSheet.Cells[f, c].Value = "EMP_TIT";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "Cliente";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CIF";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CCONTRPS";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "VER_CONTR_PS";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CCOUNIPS";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CUPS_EXT";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;


                workSheet.Cells[f, c].Value = "codSolATR";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "estadoSolATR";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "TARIFA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                for(int w = 1; w <= 6; w++)
                {
                    workSheet.Cells[f, c].Value = "POTENCIA" + w;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;
                }

                

                workSheet.Cells[f, c].Value = "fRecepcion";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "fEnvioATR";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "fAcepRech";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "fActivacion";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "fRechazo";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "MOT_RECH_1";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "COMENT_1_MOTRECH_1";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "MOT_RECH_2";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "COMENT_1_MOTRECH_2";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;                

                workSheet.Cells[f, c].Value = "tipoActivacionPrevista";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "MOD_FECHA_SOLCT";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "FechaFijaSolicitada";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                #endregion



                strSql = "SELECT s.EMP_TIT, s.cliente, s.CIF, s.CCONTRPS, s.VER_CONTR_PS,"
                    + " s.CCOUNIPS, s.CUPS_EXT, s.codSolATR, s.estadoSolATR, s.TARIFA," 
                    + " s.POTENCIA1, s.POTENCIA2, s.POTENCIA3, s.POTENCIA4, s.POTENCIA5, s.POTENCIA6,"
                    + " fRecepcion AS fechaRecepcion, fEnvioATR AS fechaEnvio, fAcepRech AS fechaAceptacion,"
                    + " fActivacion AS fechaActivacion, fRechazo AS fechaRechazo," 
                    + " s.MOT_RECH_1, s.COMENT_1_MOTRECH_1, s.MOT_RECH_2, s.COMENT_1_MOTRECH_2,"
                    + " p.descripcion AS tipoActivacionPrevista, s.MOD_FECHA_SOLCT, s.FECHA_SOLCT as FechaFijaSolicitada"
                    + " FROM cont.SOLATRMT s"
                    + " LEFT OUTER JOIN param_tipoActivacionPrevista p ON s.MOD_FECHA_RESP = p.codigo"
                    + " WHERE s.fRecepcion > 20200311 AND s.tipoSolATR = 'MODIFICACION'"
                    + " ORDER BY s.CIF, s.CCOUNIPS, s.fRecepcion;";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    hayDatos = true;

                    f++;
                    c = 1;

                    if (r["EMP_TIT"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["EMP_TIT"]);
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                     
                    c++;
                    if (r["cliente"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["cliente"].ToString();
                    c++;
                    if (r["CIF"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CIF"].ToString();
                    c++;
                    if (r["CCONTRPS"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CCONTRPS"].ToString();
                    c++;
                    if (r["VER_CONTR_PS"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["VER_CONTR_PS"]);
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                        
                    c++;
                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CCOUNIPS"].ToString();
                    c++;
                    if (r["CUPS_EXT"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CUPS_EXT"].ToString();
                    c++;
                    if (r["codSolATR"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["codSolATR"].ToString();
                    c++;
                    if (r["estadoSolATR"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["estadoSolATR"].ToString();
                    c++;
                    if (r["TARIFA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                     

                    for(int w = 1; w <= 6; w++)
                    {
                        c++;
                        if (r["POTENCIA" + w] != System.DBNull.Value)
                            if(Convert.ToInt32(r["POTENCIA" + w]) > 0)
                            {
                                workSheet.Cells[f, c].Value = Convert.ToDouble(r["POTENCIA" + w]);
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            }
                                
                    }
                    
                    c++;
                    if (r["fechaRecepcion"] != System.DBNull.Value)
                    {
                        if (Convert.ToInt32(r["fechaRecepcion"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["fechaRecepcion"].ToString().Substring(0, 4)),
                                             Convert.ToInt32(r["fechaRecepcion"].ToString().Substring(4, 2)),
                                             Convert.ToInt32(r["fechaRecepcion"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }


                    }

                    c++;
                    if (r["fechaEnvio"] != System.DBNull.Value)
                    {
                        if (Convert.ToInt32(r["fechaEnvio"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["fechaEnvio"].ToString().Substring(0, 4)),
                                             Convert.ToInt32(r["fechaEnvio"].ToString().Substring(4, 2)),
                                             Convert.ToInt32(r["fechaEnvio"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }


                    }

                    c++;
                    if (r["fechaAceptacion"] != System.DBNull.Value)
                    {
                        if (Convert.ToInt32(r["fechaAceptacion"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["fechaAceptacion"].ToString().Substring(0, 4)),
                                             Convert.ToInt32(r["fechaAceptacion"].ToString().Substring(4, 2)),
                                             Convert.ToInt32(r["fechaAceptacion"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }


                    }

                    c++;
                    if (r["fechaActivacion"] != System.DBNull.Value)
                    {
                        if (Convert.ToInt32(r["fechaActivacion"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["fechaActivacion"].ToString().Substring(0, 4)),
                                             Convert.ToInt32(r["fechaActivacion"].ToString().Substring(4, 2)),
                                             Convert.ToInt32(r["fechaActivacion"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }


                    }
                    c++;
                    if (r["fechaRechazo"] != System.DBNull.Value)
                    {
                        if(Convert.ToInt32(r["fechaRechazo"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["fechaRechazo"].ToString().Substring(0, 4)),
                                             Convert.ToInt32(r["fechaRechazo"].ToString().Substring(4, 2)),
                                             Convert.ToInt32(r["fechaRechazo"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        

                    }
                    c++;
                    if (r["MOT_RECH_1"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["MOT_RECH_1"].ToString();
                    c++;
                    if (r["COMENT_1_MOTRECH_1"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["COMENT_1_MOTRECH_1"].ToString();
                    c++;
                    if (r["MOT_RECH_2"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["MOT_RECH_2"].ToString();
                    c++;
                    if (r["COMENT_1_MOTRECH_2"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["COMENT_1_MOTRECH_2"].ToString();
                    c++;
                    if (r["tipoActivacionPrevista"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["tipoActivacionPrevista"].ToString();
                    c++;
                    if (r["MOD_FECHA_SOLCT"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["MOD_FECHA_SOLCT"].ToString();
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                        
                    c++;                    
                    if (r["FechaFijaSolicitada"] != System.DBNull.Value)
                    {
                        if (Convert.ToInt32(r["FechaFijaSolicitada"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["FechaFijaSolicitada"].ToString().Substring(0, 4)),
                                             Convert.ToInt32(r["FechaFijaSolicitada"].ToString().Substring(4, 2)),
                                             Convert.ToInt32(r["FechaFijaSolicitada"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }


                    }

                }
                db.CloseConnection();

                var allCells = workSheet.Cells[1, 1, f, 28];
                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:AB1"].AutoFilter = true;
                allCells.AutoFitColumns();

                excelPackage.Save();

                if (hayDatos)
                {
                    #region EnvioMail
                    office.SendMail s = new office.SendMail("Contratacionee@enel.com");
                    StringBuilder textBody = new StringBuilder();
                    s.para = utilidades.GOMail.Destinatarios("covid_solatrmt", utilidades.GOMail.destinatarios.to);
                    s.cc = utilidades.GOMail.Destinatarios("covid_solatrmt", utilidades.GOMail.destinatarios.cc);
                    s.asunto = "Listado solicitudes ATR a " + DateTime.Now.ToString("dd/MM/yyyy");
                    s.adjuntos.Add(fileInfo.FullName);

                    textBody.Append(System.Environment.NewLine);
                    textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("   Os actualizo el fichero con las solicitudes del circuito SE.");
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("Un saludo.");

                    s.htmlCuerpo = textBody.ToString();
                    if (p.GetValue("EnviarMail", DateTime.Now, DateTime.Now) == "S")
                        s.Send();
                    else
                        s.Save();
                    #endregion
                }
                else
                {
                    #region EnvioMail
                    office.SendMail s = new office.SendMail("Contratacionee@enel.com");
                    StringBuilder textBody = new StringBuilder();
                    s.para = utilidades.GOMail.Destinatarios("covid_solatrmt", utilidades.GOMail.destinatarios.to);
                    s.cc = utilidades.GOMail.Destinatarios("covid_solatrmt", utilidades.GOMail.destinatarios.cc);
                    s.asunto = "Listado solicitudes ATR  a " + DateTime.Now.ToString("dd/MM/yyyy");

                    textBody.Append(System.Environment.NewLine);
                    textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("   No se han detectado solicitudes del circuito SE.");
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("Un saludo.");

                    s.htmlCuerpo = textBody.ToString();
                    if (p.GetValue("EnviarMail", DateTime.Now, DateTime.Now) == "S")
                        s.Send();
                    else
                        s.Save();
                    #endregion
                }



            }
            catch (Exception e)
            {

                ficheroLog.AddError("Covid.SolatrMT_Covid19 --> " + e.Message);
                EndesaBusiness.mail.MailExchangeServer mes =
                    new EndesaBusiness.mail.MailExchangeServer(p.GetValue("usuario_errores", DateTime.Now, DateTime.Now));
                mes.SendMail(p.GetValue("errores_from", DateTime.Now, DateTime.Now),
                    p.GetValue("errores_to", DateTime.Now, DateTime.Now), null, "Error en Covid.SolatrMT_Covid19 " + DateTime.Now.ToString("dd/MM/yyyy"),
                    System.Environment.NewLine +
                    System.Environment.NewLine +
                    e.Message, null);
            }
        }


    }
}
