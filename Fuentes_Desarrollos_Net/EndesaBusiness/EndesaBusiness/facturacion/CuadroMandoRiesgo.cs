using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    
    public class CuadroMandoRiesgo
    {
        public Dictionary<string, EndesaEntity.facturacion.CuadroMando> dic { get; set; }
        utilidades.Param p;
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_cuadro_mando_riesgo");
        public CuadroMandoRiesgo()
        {


            
            FileInfo fichero;
            string prefijoArchivo = "";
            string rutaSalida = "";
            p = new utilidades.Param("cm_param", MySQLDB.Esquemas.FAC);

            
            ficheroLog.Add("===========");
            ficheroLog.Add("INICIO PROCESO");
            ficheroLog.Add("===========");
            ficheroLog.Add("");

            dic = CargaInventario();
            ficheroLog.Add("Carga de Inventario con " + dic.Count + " registros");
            
            utilidades.PdteWeb pdteWeb = 
                new utilidades.PdteWeb(dic.Where(z => z.Value.electricidad_gas == "Electricidad").Select(z => z.Value.cups13).ToList());
            ficheroLog.Add("Carga de Pdte Web");

            ficheroLog.Add("Carga de contratos ps");
            contratacion.PS_AT ps = 
                new contratacion.PS_AT(dic.Where(z => z.Value.electricidad_gas == "Electricidad").Select(z => z.Value.cupsree).ToList());

            ficheroLog.Add("Carga de TAM electricidad");
            TamElectricidad te = 
                new TamElectricidad(dic.Where(z => z.Value.electricidad_gas == "Electricidad").Select(z => z.Value.cupsree).ToList());

            ficheroLog.Add("Carga de TAM GAS");
            TamGas tg = 
                new TamGas(dic.Where(z => z.Value.electricidad_gas == "Gas").Select(z => z.Value.cupsree).ToList());

            ficheroLog.Add("Carga de historico cuadro de mando de facturacion");

            HistoricoCuadroMandoFacturacion hcm =
                new HistoricoCuadroMandoFacturacion(dic.Select(z => z.Value.cupsree).ToList());

            ficheroLog.Add("Carga puntos Gas");

            gas.PuntosActivosGas puntosGas =
                new gas.PuntosActivosGas(dic.Where(z => z.Value.electricidad_gas == "Gas").Select(z => z.Value.cupsree).ToList());

            #region Asignacion Valores
            foreach (KeyValuePair<string, EndesaEntity.facturacion.CuadroMando> p in dic)
            {

                if (hcm.Existe_CUSP20(p.Key))
                {
                    p.Value.tipo = hcm.tipo;
                    p.Value.mes = hcm.ultimo_mes_facturado.ToString();
                    p.Value.estado = hcm.estado_ltp;
                    //p.Value.estado = hcm.estado_contrato;
                }
                

                if (p.Value.electricidad_gas == "Electricidad")
                {
                    if (ps.ExisteAlta(p.Key))
                    {
                        p.Value.estado_contrato = ps.estado_contrato_descripcion;
                        p.Value.provincia = ps.provincia;
                        pdteWeb.GetEstado(p.Value.cups13);
                        p.Value.estado = pdteWeb.estado_ltp;
                    }
                    p.Value.tam = te.GetTamCups20(p.Key);
                }
                else
                {
                    p.Value.tam = tg.GetTamCups20(p.Key);
                    if (puntosGas.ExisteAlta(p.Key))
                    {
                        p.Value.estado_contrato = puntosGas.estado_contrato_descripcion;
                        p.Value.provincia = puntosGas.provincia;
                    }
                    
                }
                    
            }
            #endregion

            

            prefijoArchivo = p.GetValue("nombre_archivo_cm_riesgo", DateTime.Now, DateTime.Now);
            rutaSalida = p.GetValue("ruta_salida_cm_riesgo", DateTime.Now, DateTime.Now);
            fichero = new FileInfo(rutaSalida + prefijoArchivo + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
            ficheroLog.Add("Generación informe Excel");
            InformeExcel(fichero.FullName, dic);

            //email.SaveMail(p.GetValue("riesgo_mail_from", DateTime.Now, DateTime.Now),
            //    p.GetValue("riesgo_mail_to", DateTime.Now, DateTime.Now),
            //    p.GetValue("riesgo_mail_cc", DateTime.Now, DateTime.Now),
            //    "Cuadro de Mando Riesgo " + DateTime.Now.ToString("dd/MM/yyyy"),
            //    "Buenos días:" +
            //    System.Environment.NewLine +
            //    "   Adjunto Cuadro de Mando Riesgo." +
            //    System.Environment.NewLine +
            //    "Un saludo.",
            //    fichero.FullName);

            ficheroLog.Add("Generación Mail");
            #region EnvioMail

            //office.SendMail s = new office.SendMail(p.GetValue("Contratacionee@enel.com"));
            //StringBuilder textBody = new StringBuilder();

            //// Para pruebas
            ////s.para.Add("gabriel.mora@atos.net");
            ////s.cc.Add("gmoraarias@hotmail.com");

            //s.para = utilidades.GOMail.Destinatarios("cm_riesgo", utilidades.GOMail.destinatarios.to);
            //s.cc = utilidades.GOMail.Destinatarios("cm_riesgo", utilidades.GOMail.destinatarios.cc);
            //s.asunto = "Cuadro de Mando Riesgo " + DateTime.Now.ToString("dd/MM/yyyy");            
            //s.adjuntos.Add(fichero.FullName);

            //textBody.Append(System.Environment.NewLine);
            //textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
            //textBody.Append(System.Environment.NewLine);
            //textBody.Append("   Adjunto Cuadro de Mando Riesgo.");
            //textBody.Append(System.Environment.NewLine);
            //textBody.Append("Un saludo.");

            //s.htmlCuerpo = textBody.ToString();
            //s.Send();


            StringBuilder textBody = new StringBuilder();
            // EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
            EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();


            string from = p.GetValue("riesgo_mail_from");
            string to = utilidades.GOMail.Destinatarios_lista("cm_riesgo", utilidades.GOMail.destinatarios.to);
            string cc = utilidades.GOMail.Destinatarios_lista("cm_riesgo", utilidades.GOMail.destinatarios.cc);
            string subject = "Cuadro de Mando Riesgo " + DateTime.Now.ToString("dd/MM/yyyy");

            textBody.Append(System.Environment.NewLine);
            textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
            textBody.Append(System.Environment.NewLine);
            textBody.Append("   Adjunto Cuadro de Mando Riesgo.");
            textBody.Append(System.Environment.NewLine);
            textBody.Append("Un saludo.");

            mes.SaveMail(from, to, cc, subject, textBody.ToString(), fichero.FullName);

            ficheroLog.Add("");
            ficheroLog.Add("===========");
            ficheroLog.Add("FIN PROCESO");
            ficheroLog.Add("===========");
            #endregion


        }

        private Dictionary<string, EndesaEntity.facturacion.CuadroMando> CargaInventario()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            Dictionary<string, EndesaEntity.facturacion.CuadroMando> d = 
                new Dictionary<string, EndesaEntity.facturacion.CuadroMando>();

            try
            {
                strSql = "SELECT nif, cliente, cups13, cupsree, electricidad_gas"
                    + " FROM fact.cm_inventario_riesgo";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.CuadroMando c = new EndesaEntity.facturacion.CuadroMando();
                    c.nif = r["nif"].ToString();
                    c.cliente = r["cliente"].ToString();

                    if(r["cups13"] != System.DBNull.Value)
                        c.cups13 = r["cups13"].ToString();

                    c.cupsree = r["cupsree"].ToString();
                    c.electricidad_gas = r["electricidad_gas"].ToString();

                    EndesaEntity.facturacion.CuadroMando o;
                    if (!d.TryGetValue(c.cupsree, out o))
                        d.Add(c.cupsree, c);
                }
                db.CloseConnection();
                return d;
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        public void InformeExcel(string fichero, Dictionary<string, EndesaEntity.facturacion.CuadroMando> d)
        {
            int f = 1;
            int c = 1;            

            FileInfo fileInfo = new FileInfo(fichero);

            if (fileInfo.Exists)
                fileInfo.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Inventario Riesgo");

            var headerCells = workSheet.Cells[1, 1, 1, 28];
            var headerFont = headerCells.Style.Font;
            f = 1;

            headerFont.Bold = true;
            #region Cabecera_Excel
            
            workSheet.Cells[f, c].Value = "NIF";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Cliente";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "CUPSREE";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Electricidad/Gas";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Estado Contrato";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Provincia";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Último Periodo Facturado";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;
           

            workSheet.Cells[f, c].Value = "Estado LTP";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;
            
            workSheet.Cells[f, c].Value = "Tipo Cliente";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "TAM";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            #endregion

            foreach(KeyValuePair<string, EndesaEntity.facturacion.CuadroMando> p in d)
            {
                c = 1;
                f++;
                
                workSheet.Cells[f, c].Value = p.Value.nif; c++;
                workSheet.Cells[f, c].Value = p.Value.cliente; c++;
                workSheet.Cells[f, c].Value = p.Value.cupsree; c++;
                workSheet.Cells[f, c].Value = p.Value.electricidad_gas; c++;
                if(p.Value.estado_contrato != null)
                    workSheet.Cells[f, c].Value = p.Value.estado_contrato.ToUpper(); 
                c++;
                workSheet.Cells[f, c].Value = p.Value.provincia; c++;
                if(Convert.ToInt32(p.Value.mes) > 0)
                    workSheet.Cells[f, c].Value = p.Value.mes.ToString(); 
                c++;
                workSheet.Cells[f, c].Value = p.Value.estado; c++;
                workSheet.Cells[f, c].Value = p.Value.tipo; c++;                                
                workSheet.Cells[f, c].Value = p.Value.tam; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
            }

            var allCells = workSheet.Cells[1, 1, f, 28];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:J1"].AutoFilter = true;
            allCells.AutoFitColumns();
         
            excelPackage.Save();
            excelPackage = null;


        }

    }
}
