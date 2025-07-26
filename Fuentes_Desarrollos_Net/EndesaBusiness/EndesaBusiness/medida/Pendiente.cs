using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.utilidades;
using EndesaBusiness.servidores;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Style;
using System.Globalization;
using System.Windows.Forms;
using Org.BouncyCastle.Bcpg;
using System.Threading;

namespace EndesaBusiness.medida
{

   
    public class Pendiente : EndesaEntity.medida.Pendiente
    {
        utilidades.Param p;
        utilidades.Param param;
        utilidades.Param local_param;
        logs.Log ficheroLog;        

        Dictionary<DateTime, Dictionary<string, List<EndesaEntity.medida.Pendiente>>> dic;
        Dictionary<string, List<EndesaEntity.medida.Pendiente>> dicPendiente;
        Dictionary<string, List<EndesaEntity.medida.Pendiente>> dicPendienteNormal;        

        contratacion.ComplementosContrato complementosContratoPS;
        contratacion.Agora_Portugal agoraPortugal;

        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;

        public Pendiente()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_Pendiente");
            param = new utilidades.Param("global_param", servidores.MySQLDB.Esquemas.AUX);
            p = new utilidades.Param("ps_param", servidores.MySQLDB.Esquemas.CON);
            local_param = new utilidades.Param("dt_vw_ed_f_detalle_pendiente_facturar_param", servidores.MySQLDB.Esquemas.MED);
            ss_pp = new utilidades.Seguimiento_Procesos();

        }

        public Pendiente(List<string> lista_cups13)
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_Pendiente");
            param = new utilidades.Param("global_param", servidores.MySQLDB.Esquemas.AUX);
            p = new utilidades.Param("ps_param", servidores.MySQLDB.Esquemas.CON);
            local_param = new utilidades.Param("dt_vw_ed_f_detalle_pendiente_facturar_param", servidores.MySQLDB.Esquemas.MED);
            dicPendienteNormal = CargaPendienteNormal(lista_cups13);
            ss_pp = new utilidades.Seguimiento_Procesos();
        }

        public void CargaPendiente()
        {
            dicPendiente = CargaPendienteTotal();
        }

        public void CargaPendiente(DateTime fecha)
        {
            dicPendiente = CargaPendienteTotal(fecha);
        }


        private Dictionary<string, List<EndesaEntity.medida.Pendiente>> CargaPendienteNormal(List<string> lista_cups13)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;


            Dictionary<string, List<EndesaEntity.medida.Pendiente>> d =
                new Dictionary<string, List<EndesaEntity.medida.Pendiente>>();


            try
            {
                strSql = "SELECT empresa_titular, punto_de_medida, contrato, mes, distribuidora, estado,"
                    + "subestado, multimedida, fh_desde, fecha_informe"
                    + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar where"
                    + " substr(punto_de_medida,1,13) in ('" + lista_cups13[0] + "'";

                for (int i = 1; i < lista_cups13.Count; i++)
                    strSql += " ,'" + lista_cups13[i] + "'";

                strSql += ")";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                    c.empresaTitular = r["empresa_titular"].ToString();
                    c.cups15 = r["punto_de_medida"].ToString();
                    c.cups13 = c.cups15.Substring(0, 13);
                    c.aaaammPdte = Convert.ToInt32(r["mes"]);
                    c.estado = r["estado"].ToString();
                    c.subsEstado = r["subestado"].ToString();
                    c.multimedida = (r["multimedida"].ToString() == "S");
                    c.fh_desde = Convert.ToDateTime(r["fh_desde"]);



                    List<EndesaEntity.medida.Pendiente> o;
                    if (!d.TryGetValue(c.cups13, out o))
                    {
                        o = new List<EndesaEntity.medida.Pendiente>();
                        o.Add(c);
                        d.Add(c.cups13, o);
                    }else
                        o.Add(c);
                
                }
                db.CloseConnection();
                return d;
            }catch(Exception ex)
            {
                ficheroLog.addError("CargaPendienteNormal: " + ex.Message);
                return null;
            }


        }
        private Dictionary<string, List<EndesaEntity.medida.Pendiente>> CargaPendienteTotal()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, List<EndesaEntity.medida.Pendiente>> d
                = new Dictionary<string, List<EndesaEntity.medida.Pendiente>>();

            try
            {
                strSql = " SELECT pend.empresa_titular AS EMPRESA,"
                    + " pend.cups13, ps.CUPS20,"
                    + " pend.mes as aaaammPdte, pend.estado, pend.subestado"
                    + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar_agrupado pend"
                    + " LEFT OUTER JOIN cont.PS_AT ps ON"
                    + " ps.IDU = pend.cups13"
                    + " ORDER BY pend.empresa_titular, "
                    + " pend.cups13, pend.mes ASC";

                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                    c.empresaTitular = r["EMPRESA"].ToString();
                    c.cups13 = r["cups13"].ToString().ToUpper();
                    c.aaaammPdte = Convert.ToInt32(r["aaaammPdte"]);
                    c.estado = r["estado"].ToString();
                    c.subsEstado = r["subestado"].ToString();
                    if (r["CUPS20"] != System.DBNull.Value)
                        c.cups20 = r["CUPS20"].ToString();

                    List<EndesaEntity.medida.Pendiente> o;
                    if (!d.TryGetValue(c.cups13, out o))
                    {
                        o = new List<EndesaEntity.medida.Pendiente>();
                        o.Add(c);
                        d.Add(c.cups13, o);
                    }
                    else
                        o.Add(c);


                        
                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                ficheroLog.addError("CargaPendiente: " + e.Message);
                return null;
            }
        }
        private Dictionary<string, List<EndesaEntity.medida.Pendiente>> CargaPendienteTotal(DateTime fecha)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, List<EndesaEntity.medida.Pendiente>> d
                = new Dictionary<string, List<EndesaEntity.medida.Pendiente>>();

            try
            {
                strSql = " SELECT pend.empresa_titular AS EMPRESA,"
                    + " pend.cups13, "
                    + " pend.mes as aaaammPdte, pend.estado, pend.subestado"
                    + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar_agrupado_hist pend where "
                    + " fecha_informe = '" + fecha.ToString("yyyy-MM-dd") + "'"
                    + " ORDER BY pend.empresa_titular, "
                    + " pend.cups13, pend.mes ASC";

                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                    c.empresaTitular = r["EMPRESA"].ToString();
                    c.cups13 = r["cups13"].ToString();
                    c.aaaammPdte = Convert.ToInt32(r["aaaammPdte"]);
                    c.estado = r["estado"].ToString();
                    c.subsEstado = r["subestado"].ToString();

                    List<EndesaEntity.medida.Pendiente> o;
                    if (!d.TryGetValue(c.cups13, out o))
                    {
                        o = new List<EndesaEntity.medida.Pendiente>();
                        o.Add(c);
                        d.Add(c.cups13, o);
                    }
                    o.Add(c);
                        
                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                ficheroLog.addError("CargaPendiente: " + e.Message);
                return null;
            }
        }

       

        public void GetCUPS20(string cups20)
        {
            this.existe = false;

            foreach(KeyValuePair<string, List<EndesaEntity.medida.Pendiente>> p in dicPendiente)
            {
                if (this.existe)
                    break;

                foreach (EndesaEntity.medida.Pendiente pp in p.Value)
                {
                    if (pp.cups20 == cups20)
                    {
                        this.existe = true;
                        this.estado = pp.estado;
                        this.subsEstado = pp.subsEstado;
                        this.aaaammPdte = pp.aaaammPdte;
                        this.descripcion_estado = pp.descripcion_estado;
                        this.descripcion_subestado = pp.descripcion_subestado;
                        break;
                    }
                }
            }
                



           
        }
        public void GetCups13(string cups13)
        {
            List<EndesaEntity.medida.Pendiente> o;
            if (dicPendiente.TryGetValue(cups13, out o))
            {
                this.existe = true;
                this.estado = o[0].estado;
                this.subsEstado = o[0].subsEstado;
                this.aaaammPdte = o[0].aaaammPdte;
            }
            else
            {
                this.existe = false;
            }
        } 

        public void GetCups13_Normal(string cups13)
        {
            List<EndesaEntity.medida.Pendiente> o;
            if (dicPendienteNormal.TryGetValue(cups13, out o))
            {

                this.existe = true;
                this.cups15 = o[0].cups15;
                this.estado = o[0].estado;
                this.subsEstado = o[0].subsEstado;
                this.aaaammPdte = o[0].aaaammPdte;
                this.fh_desde = o[0].fh_desde;
                this.multimedida = o[0].multimedida;
            }
            else
            {
                this.existe = false;
            }
        }

        public void InformePend()
        {
            
            string archivo_informe = "";

            EndesaBusiness.utilidades.FechasProcesos fp = new utilidades.FechasProcesos();

            try
            {

                ficheroLog.Add("Actualizado Pendiente: " + (ActualizadoPendiente() ? "S" : "N"));
                Console.WriteLine("Actualizado Pendiente: " + (ActualizadoPendiente() ? "S" : "N"));                
                ficheroLog.Add("Última actualización Informe: " + ss_pp.GetFecha_FinProceso("Facturación", "INF_PDTE_WEB", "INF_PDTE_WEB").ToString("dd/MM/yyyy HH:mm:ss"));
                Console.WriteLine("Última actualización Informe: " + ss_pp.GetFecha_FinProceso("Facturación", "INF_PDTE_WEB", "INF_PDTE_WEB").ToString("dd/MM/yyyy HH:mm:ss"));

                if (ActualizadoPendiente() && (ss_pp.GetFecha_FinProceso("Facturación", "INF_PDTE_WEB", "INF_PDTE_WEB").Date < DateTime.Now.Date))
                {
                    ss_pp.Update_Fecha_Inicio("Facturación", "INF_PDTE_WEB", "INF_PDTE_WEB");
                    EndesaBusiness.utilidades.Fichero.BorrarArchivos_MenosUltimo(param.GetValue("fac_inf_pdte_web_ruta_salida"), param.GetValue("fac_inf_pdte_web_nombre_informe"), "xlsx", null);
                    archivo_informe = GeneraInforme();
                    if(param.GetValue("fac_inf_pdte_web_enviar_mail") == "S")
                        EnvioCorreo(archivo_informe);

                    ss_pp.Update_Fecha_Fin("Facturación", "INF_PDTE_WEB", "INF_PDTE_WEB");


                }
            }catch(Exception e)
            {
                ficheroLog.addError("InformePend: " + e.Message);
                Console.WriteLine(e.Message);
            }
            
        }

        private string GeneraInforme()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage;
            FileInfo fileInfo;            

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            int f = 1;
            int c = 1;

            utilidades.Fechas utilfecha = new Fechas();
            bool firstOnly = true;
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_totales;
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_agora;
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_Noagora;

            DateTime fd = new DateTime();

            try
            {

                fd = utilfecha.UltimoDiaHabil();
                fd = fd.AddDays(-31);

                ficheroLog.Add("Generando Informe");
                ficheroLog.Add("=================");

                fileInfo = new FileInfo(param.GetValue("fac_inf_pdte_web_ruta_salida")
                + param.GetValue("fac_inf_pdte_web_nombre_informe")
                + DateTime.Now.ToString("yyyy_MM_dd_HHmmss") + ".xlsx");

                dic_totales = CargaTotales();
                dic_agora = CargaAgora(fd);
                dic_Noagora = CargaNoAgora(fd);

                excelPackage = new ExcelPackage(fileInfo);

                var workSheet = excelPackage.Workbook.Worksheets.Add("Totales");

                var headerCells = workSheet.Cells[1, 1, 1, 8];
                var headerFont = headerCells.Style.Font;

                var allCells = workSheet.Cells[1, 1, 1, 8];
                var cellFont = allCells.Style.Font;

                #region cabecera

                f++;

                workSheet.Cells[f, c].Value = "PENDIENTE TOTAL";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); f++;

                workSheet.Cells[f, c].Value = "01. Pendiente de medida";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;
                

                workSheet.Cells[f, c].Value = "    " + "01.A. Endesa - Telemedida";                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;
               

                workSheet.Cells[f, c].Value = "    " + "01.B. Endesa - TPLa";                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;
              

                workSheet.Cells[f, c].Value = "    " + "01.D. NoEndesa - Telemedidaa";                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;
             

                workSheet.Cells[f, c].Value = "    " + "01.E. NoEndesa - TPL";
                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "02. CC Rechazada por CS";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;
                

                workSheet.Cells[f, c].Value = "    " + "02.A. El Punto no existe en el CS";
                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "03. CC Completa en CS";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;
                

                workSheet.Cells[f, c].Value = "    " + "03.A. CC Completa en el CS";
                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "04. CC Enviada a SCE ML";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;
                

                workSheet.Cells[f, c].Value = "    " + "04.A. CC Enviada a SCE ML";                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "06. CC Incompleta  SCE ML";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;
                

                workSheet.Cells[f, c].Value = "    " + "06.A. CC Incompleta  SCE ML";
                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "07. LTP SCE";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "07.A. No Validada";
                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;
                
                workSheet.Cells[f, c].Value = "    " + "07.B. Validada con Anomalías";
                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;
                           
                workSheet.Cells[f, c].Value = "    " + "07.C. Validada sin Anomalías";
                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "07.E. Modificada";
                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "08. El Punto no está Extraído";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;
                
                workSheet.Cells[f, c].Value = "    " + "08.A. Facturado en el día anterior";
                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;                

                workSheet.Cells[f, c].Value = "    " + "08.B. Bloqueado";
                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;               

                workSheet.Cells[f, c].Value = "    " + "08.C. Otros";
                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "09. El Punto está Extraído";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "09.A. El Punto está Extraído";
                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "10. Prefactura pendiente";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "10.A. Prefactura pendiente";                
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "Total general";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); f++;

                f++;
                

                workSheet.Cells[f, c].Value = "PENDIENTE ÁGORA";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); f++;

                workSheet.Cells[f, c].Value = "01. Pendiente de medida";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "01.A. Endesa - Telemedida";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "01.B. Endesa - TPLa";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "01.D. NoEndesa - Telemedidaa";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "01.E. NoEndesa - TPL";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "02. CC Rechazada por CS";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "02.A. El Punto no existe en el CS";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "03. CC Completa en CS";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "03.A. CC Completa en el CS";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "04. CC Enviada a SCE ML";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "04.A. CC Enviada a SCE ML";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "06. CC Incompleta  SCE ML";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "06.A. CC Incompleta  SCE ML";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "07. LTP SCE";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "07.A. No Validada";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "07.B. Validada con Anomalías";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "07.C. Validada sin Anomalías";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "07.E. Modificada";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "08. El Punto no está Extraído";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "08.A. Facturado en el día anterior";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "08.B. Bloqueado";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "08.C. Otros";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "09. El Punto está Extraído";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "09.A. El Punto está Extraído";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "10. Prefactura pendiente";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "10.A. Prefactura pendiente";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "Total general";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); f++;

                f++;               

                workSheet.Cells[f, c].Value = "PENDIENTE NO ÁGORA";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); f++;

                workSheet.Cells[f, c].Value = "01. Pendiente de medida";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "01.A. Endesa - Telemedida";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "01.B. Endesa - TPLa";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "01.D. NoEndesa - Telemedidaa";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "01.E. NoEndesa - TPL";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "02. CC Rechazada por CS";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "02.A. El Punto no existe en el CS";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "03. CC Completa en CS";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "03.A. CC Completa en el CS";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "04. CC Enviada a SCE ML";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "04.A. CC Enviada a SCE ML";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "06. CC Incompleta  SCE ML";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;


                workSheet.Cells[f, c].Value = "    " + "06.A. CC Incompleta  SCE ML";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "07. LTP SCE";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "07.A. No Validada";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "07.B. Validada con Anomalías";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "07.C. Validada sin Anomalías";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "07.E. Modificada";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "08. El Punto no está Extraído";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "08.A. Facturado en el día anterior";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "08.B. Bloqueado";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "08.C. Otros";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "09. El Punto está Extraído";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "09.A. El Punto está Extraído";

                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "10. Prefactura pendiente";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "    " + "10.A. Prefactura pendiente";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); f++;

                workSheet.Cells[f, c].Value = "Total general";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aquamarine); f++;

                

                #endregion

                int totales_dia = 0;
                foreach(KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> p in dic_totales)
                {

                    totales_dia = 0;
                    for (int i = 0; i < p.Value.Count; i++)
                        totales_dia = totales_dia + p.Value[i].num_cups;

                    #region diferencia
                    if (firstOnly)
                    {
                        firstOnly = false;


                        f = 2;
                        c++;

                        workSheet.Cells[f, c].Value = "DIF (D-C)";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.B. Endesa - TPL");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "02. CC Rechazada por CS", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "03. CC Completa en CS", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "03. CC Completa en CS", "03.A. CC Completa en el CS");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "04. CC Enviada a SCE ML", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "06. CC Incompleta  SCE ML", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.A. No Validada");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.B. Validada con Anomalías");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.C. Validada sin Anomalías");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.E. Modificada");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", "08.B. Bloqueado");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", "08.C. Otros");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "09. El Punto está Extraído", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "10. Prefactura pendiente", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pdte_DIA(dic_totales, utilfecha.UltimoDiaHabilAnterior(p.Key)) - totales_dia;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, c].Style.Font.Bold = true;

                    }
                    #endregion

                    c++;
                    
                    

                    f = 2;

                    workSheet.Cells[f, c].Value = p.Key;
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "01. Pendiente de medida", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "01. Pendiente de medida", "01.A. Endesa - Telemedida");                    
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "01. Pendiente de medida", "01.B. Endesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "02. CC Rechazada por CS", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "03. CC Completa en CS", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "03. CC Completa en CS", "03.A. CC Completa en el CS");                    
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "04. CC Enviada a SCE ML", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");                    
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "06. CC Incompleta  SCE ML", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");                    
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "07. LTP SCE", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "07. LTP SCE", "07.A. No Validada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "07. LTP SCE", "07.B. Validada con Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "07. LTP SCE", "07.C. Validada sin Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "07. LTP SCE", "07.E. Modificada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "08. El Punto no está Extraído", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "08. El Punto no está Extraído", "08.B. Bloqueado");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "08. El Punto no está Extraído", "08.C. Otros");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "09. El Punto está Extraído", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "10. Prefactura pendiente", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_totales, p.Key, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                                        
                    workSheet.Cells[f, c].Value = totales_dia;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;

                    

                }

                c = 1;
                firstOnly = true;

                foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> p in dic_agora)
                {

                    totales_dia = 0;
                    for (int i = 0; i < p.Value.Count; i++)
                        totales_dia = totales_dia + p.Value[i].num_cups;

                    #region diferencia
                    if (firstOnly)
                    {
                        firstOnly = false;


                        f = 31;
                        c++;

                        workSheet.Cells[f, c].Value = "DIF (D-C)";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.B. Endesa - TPL");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "02. CC Rechazada por CS", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "03. CC Completa en CS", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "03. CC Completa en CS", "03.A. CC Completa en el CS");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "04. CC Enviada a SCE ML", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "06. CC Incompleta  SCE ML", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.A. No Validada");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.B. Validada con Anomalías");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.C. Validada sin Anomalías");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.E. Modificada");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", "08.B. Bloqueado");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", "08.C. Otros");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "09. El Punto está Extraído", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "10. Prefactura pendiente", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pdte_DIA(dic_agora, utilfecha.UltimoDiaHabilAnterior(p.Key)) - totales_dia;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, c].Style.Font.Bold = true;

                    }
                    #endregion

                    c++;

                    f = 31;

                    workSheet.Cells[f, c].Value = p.Key;
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", "01.B. Endesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "02. CC Rechazada por CS", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "03. CC Completa en CS", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "04. CC Enviada a SCE ML", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "06. CC Incompleta  SCE ML", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", "07.A. No Validada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", "07.B. Validada con Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", "07.C. Validada sin Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", "07.E. Modificada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "08. El Punto no está Extraído", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "08. El Punto no está Extraído", "08.B. Bloqueado");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "08. El Punto no está Extraído", "08.C. Otros");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "09. El Punto está Extraído", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "10. Prefactura pendiente", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = totales_dia;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;



                }

                c = 1;
                firstOnly = true;

                foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> p in dic_Noagora)
                {

                    totales_dia = 0;
                    for (int i = 0; i < p.Value.Count; i++)
                        totales_dia = totales_dia + p.Value[i].num_cups;

                    #region diferencia
                    if (firstOnly)
                    {
                        firstOnly = false;


                        f = 60;
                        c++;

                        workSheet.Cells[f, c].Value = "DIF (D-C)";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.B. Endesa - TPL");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "02. CC Rechazada por CS", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "03. CC Completa en CS", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "03. CC Completa en CS", "03.A. CC Completa en el CS");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "04. CC Enviada a SCE ML", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "06. CC Incompleta  SCE ML", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.A. No Validada");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.B. Validada con Anomalías");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.C. Validada sin Anomalías");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.E. Modificada");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", "08.B. Bloqueado");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", "08.C. Otros");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "09. El Punto está Extraído", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "10. Prefactura pendiente", null);
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                        workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                        workSheet.Cells[f, c].Value = Total_Pdte_DIA(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(p.Key)) - totales_dia;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, c].Style.Font.Bold = true;

                    }
                    #endregion

                    c++;



                    f = 60;

                    workSheet.Cells[f, c].Value = p.Key;
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", "01.B. Endesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "02. CC Rechazada por CS", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "03. CC Completa en CS", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "04. CC Enviada a SCE ML", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "06. CC Incompleta  SCE ML", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", "07.A. No Validada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", "07.B. Validada con Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", "07.C. Validada sin Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", "07.E. Modificada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "08. El Punto no está Extraído", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "08. El Punto no está Extraído", "08.B. Bloqueado");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "08. El Punto no está Extraído", "08.C. Otros");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "09. El Punto está Extraído", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "10. Prefactura pendiente", null);
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = totales_dia;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;


                }


                allCells = workSheet.Cells[1, 1, f, 32];
                allCells.AutoFitColumns();

                #region Hoja Pdte Web

                workSheet = excelPackage.Workbook.Worksheets.Add("Pdte Web");

                headerCells = workSheet.Cells[1, 1, 1, 8];
                headerFont = headerCells.Style.Font;

                allCells = workSheet.Cells[1, 1, 1, 8];
                cellFont = allCells.Style.Font;
                cellFont.Bold = true;

                f = 1;
                c = 1;

                workSheet.Cells[f, c].Value = "EMPRESA";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "NIF";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "CLIENTE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "CUPS13";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "CUPSREE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "aaaammPdte";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Estado";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Subestado";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;


               

                strSql = " SELECT pend.empresa_titular AS EMPRESA,"
                    + " if(ps.CNIFDNIC is NULL, psat.NIF, trim(ps.CNIFDNIC)) AS NIF, "
                    + " if(ps.DAPERSOC IS NULL, psat.Cliente, ps.DAPERSOC) AS Cliente,"
                    + " substr(pend.punto_de_medida, 1, 13) AS CUPS13, "
                    + " if(ps.CUPSREE IS NULL, psat.CUPS22, ps.CUPSREE) AS CUPSREE,"
                    + " pend.mes as aaaammPdte, pend.estado, pend.subestado"
                    + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar pend"
                    + " LEFT OUTER JOIN cont.ps_total ps ON"
                    + " ps.IDU = substr(pend.punto_de_medida, 1, 13)"
                    + " LEFT OUTER JOIN cont.PS_AT psat ON"
                    + " psat.IDU = substr(pend.punto_de_medida, 1, 13)"
                    + " GROUP BY substr(pend.punto_de_medida,1,13)"
                    + " ORDER BY pend.empresa_titular, psat.NIF, trim(ps.CNIFDNIC), "
                    + " substr(pend.punto_de_medida,1,13), pend.mes ASC";

                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    c = 1;
                    f++;

                    Console.CursorLeft = 0;
                    Console.Write(f);

                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["EMPRESA"]); c++;
                    workSheet.Cells[f, c].Value = r["NIF"].ToString(); c++;
                    workSheet.Cells[f, c].Value = r["Cliente"].ToString(); c++;
                    workSheet.Cells[f, c].Value = r["CUPS13"].ToString(); c++;
                    workSheet.Cells[f, c].Value = r["CUPSREE"].ToString(); c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["aaaammPdte"]); c++;
                    workSheet.Cells[f, c].Value = r["estado"].ToString(); c++;
                    workSheet.Cells[f, c].Value = r["subestado"].ToString(); c++;
                }
                db.CloseConnection();

                allCells = workSheet.Cells[1, 1, f, 8];

                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:H1"].AutoFilter = true;
                allCells.AutoFitColumns();

                #endregion

                #region Hoja PDTE_TOTAL
                workSheet = excelPackage.Workbook.Worksheets.Add("PDTE_TOTAL");

                f = 1;
                c = 1;
                workSheet.View.FreezePanes(2, 1);
                headerFont.Bold = true;
                workSheet.Cells[f, c].Value = "EMPRESA"; c++;                
                workSheet.Cells[f, c].Value = "CUPS15"; c++;                
                workSheet.Cells[f, c].Value = "CONTRATO"; c++;
                workSheet.Cells[f, c].Value = "MES"; c++;
                workSheet.Cells[f, c].Value = "DISTRIBUIDORA"; c++;
                workSheet.Cells[f, c].Value = "ESTADO"; c++;
                workSheet.Cells[f, c].Value = "SUBESTADO"; c++;
                workSheet.Cells[f, c].Value = "MULTIMEDIDA"; c++;                

                for (int i = 1; i <= c; i++)
                {
                    workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }


                
                strSql = "SELECT empresa_titular, punto_de_medida, contrato, mes, distribuidora, estado," 
                    + " subestado, multimedida, fh_desde, fecha_informe"
                    + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar";

                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    f++;
                    c = 1;
                    if (r["empresa_titular"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["empresa_titular"]);
                    c++;
                    
                    if (r["punto_de_medida"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["punto_de_medida"].ToString();
                    c++;
                    
                    if (r["contrato"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt64(r["contrato"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = "###0";
                    }

                    c++;
                    if (r["mes"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["mes"]);
                    c++;
                    if (r["distribuidora"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["distribuidora"].ToString();
                    c++;
                    if (r["estado"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["estado"].ToString();
                    c++;
                    if (r["subestado"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["subestado"].ToString();
                    c++;
                    if (r["multimedida"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["multimedida"].ToString();
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    

                }

                headerCells = workSheet.Cells[1, 1, 1, c];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;
                allCells = workSheet.Cells[1, 1, f, c];

                allCells.AutoFitColumns();

                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:H1"].AutoFilter = true;
                allCells.AutoFitColumns();

                db.CloseConnection();


                #endregion

                excelPackage.Save();
                 Console.WriteLine("Excel Generado");

                return fileInfo.FullName;

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("GeneraInforme: " + e.Message);
                return null;
            }
        }

        private void EnvioCorreo(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                string from = param.GetValue("fac_inf_pdte_web_buzon_envio");
                string to = param.GetValue("fac_inf_pdte_web_mail_para");
                string cc = param.GetValue("fac_inf_pdte_web_mail_cc");
                string subject = param.GetValue("fac_inf_pdte_web_mail_asunto") + " " + DateTime.Now.ToString("dd/MM/yyyy");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Se adjunta el archivo ").Append(fileInfo.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");

                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (param.GetValue("fac_inf_pdte_web_enviar_mail") == "S")
                    mes.SendMail(from, to, cc, subject, textBody.ToString(), archivo);

                else
                    mes.SaveMail(from, to, cc, subject, textBody.ToString(), archivo);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("fac_inf_pdte_web_buzon_envio"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }

        private void EnvioCorreo_PdteWeb_PS_AT_TAM(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                string from = local_param.GetValue("mail_from");
                string to = local_param.GetValue("mail_to");
                string cc = local_param.GetValue("mail_cc");
                string subject = local_param.GetValue("mail_subject") + " a " + DateTime.Now.ToString("dd/MM/yyyy");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Se adjunta el archivo ").Append(fileInfo.Name).Append(".");                
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");

                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (local_param.GetValue("mail_enviar") == "S")
                    mes.SendMail(from, to, cc, subject, textBody.ToString(), archivo);

                else
                    mes.SaveMail(from, to, cc, subject, textBody.ToString(), archivo);

                ficheroLog.Add("Correo enviado desde: " + local_param.GetValue("mail_from"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }

        private void EnvioCorreo_PdteWeb_PS_AT_TAM_v2(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                string from = local_param.GetValue("mail_from_psat_tam");
                string to = local_param.GetValue("mail_to_psat_tam");
                string cc = local_param.GetValue("mail_cc_psat_tam");
                string subject = local_param.GetValue("mail_subject_psat_tam") + " a " + DateTime.Now.ToString("dd/MM/yyyy");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Se adjunta el archivo ").Append(fileInfo.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");

                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (local_param.GetValue("mail_enviar_mail_psat_tam") == "S")
                    mes.SendMail(from, to, cc, subject, textBody.ToString(), archivo);

                else
                    mes.SaveMail(from, to, cc, subject, textBody.ToString(), archivo);

                ficheroLog.Add("Correo enviado desde: " + local_param.GetValue("mail_from"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }

        private bool ActualizadoPendiente()
        {            
            DateTime fecha = new DateTime();

            try
            {

                fecha = Convert.ToDateTime(local_param.GetValue("ultima_actualizacion_copia"));                

                if (fecha.Date == DateTime.Now.Date)
                    return true;
                else
                    return false;

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        public bool ExisteFechaInformePendiente(DateTime fecha_informe)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            bool existe = false;

            strSql = "SELECT fecha_informe FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_hist"
                + " WHERE fecha_informe = '" + fecha_informe.ToString("yyyy-MM-dd") + "' LIMIT 1";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                existe = true;
            }
            db.CloseConnection();


            return existe;


        }

        public void Copia_Detalle_Pendiente_Facturar()
        {
            Int32 i = 0;
            Int32 k = 0;
            Int32 totalRegistros = 0;
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            bool firstOnly = true;
            string strSql = "";
            StringBuilder sb = new StringBuilder();

            servidores.MySQLDB mdb;
            MySqlCommand mcommand;
            DateTime ultimaFechaCopia = new DateTime();
            DateTime fechaInforme = new DateTime();
            utilidades.Fechas utilFechas = new Fechas();

            

            try
            {

                ultimaFechaCopia = Convert.ToDateTime(local_param.GetValue("ultima_actualizacion_copia"));
                ficheroLog.Add("Ultima copia: " + ultimaFechaCopia.ToString("dd/MM/yyyy HH:mm:ss"));
                Console.WriteLine("Ultima copia: " + ultimaFechaCopia.ToString("dd/MM/yyyy HH:mm:ss"));

                if (ultimaFechaCopia.Date < DateTime.Now.Date)
                {

                     ss_pp.Update_Fecha_Inicio("Medida", "PENDIENTE", "COPIA PENDIENTE REDSHIFT");

                    // Averiguamos si la fecha de la tabla del DataLake es la del último día habil

                    strSql = "select max(fecha_informe) as fecha_informe from ed_owner.vw_ed_f_detalle_pendiente_facturar";
                    ficheroLog.Add(strSql);
                    Console.WriteLine(strSql);
                    db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                    command = new OdbcCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {

                        fechaInforme = Convert.ToDateTime(r["fecha_informe"]);
                        ficheroLog.Add(fechaInforme.ToString("dd/MM/yyyy HH:mm:ss"));
                        Console.WriteLine(fechaInforme.ToString("dd/MM/yyyy HH:mm:ss"));

                        
                    }

                    db.CloseConnection();

                    if (fechaInforme.Date < utilFechas.UltimoDiaHabil().Date)
                    {

                        // Implementar aviso que el listado no esta actualizado
                        ss_pp.Update_Comentario("Medida", "PENDIENTE", "COPIA PENDIENTE REDSHIFT", "No actualizado pendiente SCE ML en RedShift");
                    }
                    else
                    {

                        #region Inicio copia 

                        strSql = "replace into dt_vw_ed_f_detalle_pendiente_facturar"
                            + " select * from dt_vw_ed_f_detalle_pendiente_facturar_hist";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        mcommand = new MySqlCommand(strSql, mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();


                        strSql = "delete from dt_vw_ed_f_detalle_pendiente_facturar";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        mcommand = new MySqlCommand(strSql, mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();


                        strSql = "select count(*) total from ed_owner.vw_ed_f_detalle_pendiente_facturar";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                        command = new OdbcCommand(strSql, db.con);
                        r = command.ExecuteReader();
                        while (r.Read())
                        {
                            if (r["total"] != System.DBNull.Value)
                            {
                                Console.WriteLine("Encontrados: " + Convert.ToInt32(r["total"]) + " registros");
                                totalRegistros = Convert.ToInt32(r["total"]);
                            }
                        }
                        db.CloseConnection();


                        strSql = "select empresa_titular, punto_de_medida, contrato, mes, distribuidora,"
                            + "estado, subestado, multimedida, fh_desde, fecha_informe"
                            + " from ed_owner.vw_ed_f_detalle_pendiente_facturar";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                        command = new OdbcCommand(strSql, db.con);
                        r = command.ExecuteReader();
                        while (r.Read())
                        {
                            i++;
                            k++;
                            if (firstOnly)
                            {
                                firstOnly = false;
                                sb.Append("replace into dt_vw_ed_f_detalle_pendiente_facturar");
                                sb.Append(" (empresa_titular,punto_de_medida,contrato,mes,distribuidora,estado,subestado,");
                                sb.Append("multimedida,fh_desde,fecha_informe) values");
                            }

                            #region Campos
                            if (r["empresa_titular"] != System.DBNull.Value)
                                sb.Append(" (").Append(Convert.ToInt32(r["empresa_titular"])).Append(",");
                            else
                                sb.Append("(null,");

                            if (r["punto_de_medida"] != System.DBNull.Value)
                                sb.Append("'").Append(r["punto_de_medida"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["contrato"] != System.DBNull.Value)
                                sb.Append(r["contrato"].ToString()).Append(",");
                            else
                                sb.Append("null,");

                            if (r["mes"] != System.DBNull.Value)
                                sb.Append(Convert.ToInt32(r["mes"])).Append(",");
                            else
                                sb.Append("null,");

                            if (r["distribuidora"] != System.DBNull.Value)
                                sb.Append("'").Append(FuncionesTexto.RT(r["distribuidora"].ToString())).Append("',");
                            else
                                sb.Append("null,");

                            if (r["estado"] != System.DBNull.Value)
                                sb.Append("'").Append(r["estado"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["subestado"] != System.DBNull.Value)
                                sb.Append("'").Append(r["subestado"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["multimedida"] != System.DBNull.Value)
                                sb.Append("'").Append(r["multimedida"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["fh_desde"] != System.DBNull.Value)
                                sb.Append("'").Append(Convert.ToDateTime(r["fh_desde"]).ToString("yyyy-MM-dd")).Append("',");
                            else
                                sb.Append("null,");

                            if (r["fecha_informe"] != System.DBNull.Value)
                                sb.Append("'").Append(Convert.ToDateTime(r["fecha_informe"]).ToString("yyyy-MM-dd")).Append("'),");
                            else
                                sb.Append("null),");

                            #endregion

                            if (i == 500)
                            {
                                Console.WriteLine("Anexamos " + k + " / " + totalRegistros);
                                firstOnly = true;
                                mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                                mcommand = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), mdb.con);
                                mcommand.ExecuteNonQuery();
                                mdb.CloseConnection();
                                sb = null;
                                sb = new StringBuilder();
                                i = 0;
                            }

                        }

                        if (i > 0)
                        {
                            Console.WriteLine("Anexamos " + k + " / " + totalRegistros);
                            firstOnly = true;
                            mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                            mcommand = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), mdb.con);
                            mcommand.ExecuteNonQuery();
                            mdb.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            i = 0;
                        }

                        strSql = "DELETE FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        mcommand = new MySqlCommand(strSql, mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();

                        strSql = "REPLACE INTO  dt_vw_ed_f_detalle_pendiente_facturar_agrupado"
                            + " SELECT empresa_titular,substr(punto_de_medida, 1, 13) AS cups13, contrato, mes, distribuidora,"
                            + " estado, subestado, multimedida, fecha_informe"
                            + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar p ORDER BY punto_de_medida, mes DESC";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        mcommand = new MySqlCommand(strSql, mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();

                        strSql = "REPLACE INTO dt_vw_ed_f_detalle_pendiente_facturar_agrupado_t"
                            + " SELECT p.*, t.TAM FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado p"
                            + " LEFT OUTER JOIN tam t ON"
                            + " t.CUPS13 = p.cups13";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        mcommand = new MySqlCommand(strSql, mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();

                        strSql = "REPLACE INTO  dt_vw_ed_f_detalle_pendiente_facturar_agrupado_hist"
                            + " SELECT * FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        mcommand = new MySqlCommand(strSql, mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();

                        strSql = "REPLACE INTO  dt_vw_ed_f_detalle_pendiente_facturar_agrupado_hist_t"
                            + " SELECT * FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_t";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        mcommand = new MySqlCommand(strSql, mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();

                        Crea_Pdte_Agora();

                        // Para el informe de totales
                        strSql = "REPLACE INTO dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales"
                            + " SELECT  fecha_informe, h.estado, h.subestado, COUNT(*) AS num_cups"
                            + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado h"
                            + " GROUP BY h.estado, h.subestado";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        mcommand = new MySqlCommand(strSql, mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();

                        strSql = "REPLACE INTO dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales_t"
                            + " SELECT  h.fecha_informe, h.empresa_titular, h.estado, h.subestado, COUNT(*) AS num_cups, SUM(tam) AS tam"
                            + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_t h"
                            + " GROUP BY  h.fecha_informe, h.empresa_titular, h.estado, h.subestado";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        mcommand = new MySqlCommand(strSql, mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();

                        strSql = "REPLACE INTO dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales_agora"
                              + " SELECT fecha_informe, h.estado, h.subestado, COUNT(*) AS num_cups"
                              + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_agora h"
                              + " WHERE h.agora = 'S'"
                              + " GROUP BY h.fecha_informe, h.estado, h.subestado";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        mcommand = new MySqlCommand(strSql, mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();

                        strSql = "REPLACE INTO dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales_agora_t"
                              + " SELECT fecha_informe, h.empresa_titular, h.estado, h.subestado, COUNT(*) AS num_cups, SUM(tam) AS tam"
                              + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_agora_t h"
                              + " WHERE h.agora = 'S'"
                              + " GROUP BY fecha_informe, h.empresa_titular, h.estado, h.subestado";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        mcommand = new MySqlCommand(strSql, mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();

                        strSql = "REPLACE INTO dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales_noagora"
                             + " SELECT fecha_informe, h.estado, h.subestado, COUNT(*) AS num_cups"
                             + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_agora h"
                             + " WHERE h.agora = 'N'"
                             + " GROUP BY h.estado, h.subestado";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        mcommand = new MySqlCommand(strSql, mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();

                        strSql = "REPLACE INTO dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales_noagora_t"
                             + " SELECT fecha_informe, h.empresa_titular, h.estado, h.subestado, COUNT(*) AS num_cups, SUM(tam) AS tam"
                             + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_agora_t h"
                             + " WHERE h.agora = 'N'"
                             + " GROUP BY h.empresa_titular, h.estado, h.subestado";
                        ficheroLog.Add(strSql);
                        Console.WriteLine(strSql);
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        mcommand = new MySqlCommand(strSql, mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();
                        #endregion

                        UpdateFechaCopia();

                        ss_pp.Update_Fecha_Fin("Medida", "PENDIENTE", "COPIA PENDIENTE REDSHIFT");
                    }
                }

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                ss_pp.Update_Comentario("Medida", "PENDIENTE", "COPIA PENDIENTE REDSHIFT", e.Message);
            }
        }

        public void Copia_Detalle_Pendiente_Facturar_Cuadro_Mando_Facturacion()
        {
            bool existe_pendiente = false;
            
            Int32 i = 0;
            Int32 k = 0;
            Int32 totalRegistros = 0;

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;


            bool firstOnly = true;
            string strSql = "";
            StringBuilder sb = new StringBuilder();

            servidores.MySQLDB db2;
            MySqlCommand command2;
            
            
            utilidades.Fechas utilFechas = new Fechas();

            utilidades.Fechas f = new utilidades.Fechas();

            try
            {


                DateTime fecha = f.UltimoDiaHabil();
                
                // Averiguamos si la fecha de la tabla del DataLake es la del último día habil
                existe_pendiente = this.ExisteFechaInformePendiente(fecha);
                if (existe_pendiente)
                {

                    fecha = new DateTime(fecha.Year, fecha.Month, 1);
                    // Actualizamos el histórico
                    strSql = "replace into cm_pendiente_ml_hist select * from cm_pendiente_ml";
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    // Únicamente nos quedamo con el pdte para el informe
                    strSql = "DELETE FROM cm_pendiente_ml WHERE F_ULT_MOD < '" + fecha.ToString("yyyy-MM-dd") + "'";
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();


                    strSql = "SELECT if(p.empresa_titular = 20, 'Endesa Energía',"
                        + " if (p.empresa_titular = 70, 'Endesa Energía XXI','')) EmpresaTitular,"
                        + " punto_de_medida AS PM, contrato AS CodContrato,"
                        + " mes AS aaaammPdte, distribuidora AS Distribuidora, estado AS Estado,"
                        + " subestado AS Subestado, fecha_informe AS F_ULT_MOD"
                        + " FROM dt_vw_ed_f_detalle_pendiente_facturar p";
                    ficheroLog.Add(strSql);
                    Console.WriteLine(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        i++;
                        k++;
                        if (firstOnly)
                        {
                            firstOnly = false;
                            sb.Append("replace into cm_pendiente_ml");
                            sb.Append(" (EmpresaTitular,PM,CodContrato,aaaammPdte,Distribuidora,Estado,Subestado,");
                            sb.Append("F_ULT_MOD) values");
                        }

                        #region Campos
                        if (r["EmpresaTitular"] != System.DBNull.Value)
                            sb.Append(" ('").Append(r["EmpresaTitular"].ToString()).Append("',");
                        else
                            sb.Append("(null,");

                        if (r["PM"] != System.DBNull.Value)
                            sb.Append("'").Append(r["PM"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["CodContrato"] != System.DBNull.Value)
                            sb.Append("'").Append(r["CodContrato"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["aaaammPdte"] != System.DBNull.Value)
                            sb.Append(Convert.ToInt32(r["aaaammPdte"])).Append(",");
                        else
                            sb.Append("null,");

                        if (r["Distribuidora"] != System.DBNull.Value)
                            sb.Append("'").Append(FuncionesTexto.RT(r["Distribuidora"].ToString())).Append("',");
                        else
                            sb.Append("null,");

                        if (r["Estado"] != System.DBNull.Value)
                            sb.Append("'").Append(r["Estado"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["Subestado"] != System.DBNull.Value)
                            sb.Append("'").Append(r["Subestado"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["F_ULT_MOD"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["F_ULT_MOD"]).ToString("yyyy-MM-dd")).Append("'),");
                        else
                            sb.Append("null),");

                        #endregion

                        if (i == 500)
                        {
                            Console.WriteLine("Anexamos " + k + " / " + totalRegistros);
                            firstOnly = true;
                            db2 = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                            command2 = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db2.con);
                            command2.ExecuteNonQuery();
                            db2.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            i = 0;
                        }

                    }

                    if (i > 0)
                    {
                        Console.WriteLine("Anexamos " + k + " / " + totalRegistros);
                        firstOnly = true;
                        db2 = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                        command2 = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db2.con);
                        command2.ExecuteNonQuery();
                        db2.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        i = 0;
                    }


                    this.ActualizaFechaProceso_OK("PdteWeb", DateTime.Now);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void ActualizaFechaProceso_OK(string proceso, DateTime d)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            ficheroLog.Add("Actualizamos la fecha del proceso " + proceso + " de la tabla cm_fechas_procesos con valor " + d.ToString("yyyy-MM-dd"));
            strSql = "replace cm_fechas_procesos"
            + " set fecha = '" + d.ToString("yyyy-MM-dd") + "', "
             + " proceso = '" + proceso + "';";


            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        private void UpdateFechaCopia()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update dt_vw_ed_f_detalle_pendiente_facturar_param"
                + " set value = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                + " where code = 'ultima_actualizacion_copia'";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
        }

        public void GeneraInformePendWeb_PSAT_TAM(bool automatico)
        {
            FileInfo file;
            string ruta_salida_archivo = "";

            string[] listaArchivos = Directory.GetFiles(local_param.GetValue("ruta_salida_informe"),
                    local_param.GetValue("prefijo_informe") + "*.xlsx");
            for (int i = 0; i < listaArchivos.Length; i++)
            {
                file = new FileInfo(listaArchivos[i]);
                file.Delete();
            }

            ruta_salida_archivo = local_param.GetValue("ruta_salida_informe")
            + local_param.GetValue("prefijo_informe")
            + DateTime.Now.ToString("yyyyMMdd")
            + local_param.GetValue("sufijo_informe");

            ss_pp.Update_Fecha_Inicio("Facturación", "Informe PDTE + PSAT + TAM", "Informe PDTE + PSAT + TAM");

            InformePendWeb_PSAT_TAM(ruta_salida_archivo, true);

            ss_pp.Update_Fecha_Fin("Facturación", "Informe PDTE + PSAT + TAM", "Informe PDTE + PSAT + TAM");
        }

        public void GeneraInformePendWeb_PSAT_TAM_V2(bool automatico)
        {
            FileInfo file;
            string ruta_salida_archivo = "";

            string[] listaArchivos = Directory.GetFiles(local_param.GetValue("ruta_salida_informe"),
                    local_param.GetValue("prefijo_informe") + "*.xlsx");
            for (int i = 0; i < listaArchivos.Length; i++)
            {
                file = new FileInfo(listaArchivos[i]);
                file.Delete();
            }

            ruta_salida_archivo = local_param.GetValue("ruta_salida_informe")
            + local_param.GetValue("prefijo_informe")
            + DateTime.Now.ToString("yyyyMMdd")
            + local_param.GetValue("sufijo_informe");

            InformePendWeb_PSAT_TAM_V2(ruta_salida_archivo, automatico);
        }

        public void InformePendWeb_PSAT_TAM_V2(string ruta_salida_archivo, bool automatico)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";


            int c = 1;
            int f = 1;


            DateTime fecha_actual = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime fecha_registro = new DateTime();
            int meses_pdtes = 0;
            int aniomes = 0;

            EndesaBusiness.facturacion.cuadro_mando.Sofisticados sof =
                new facturacion.cuadro_mando.Sofisticados();

            Dictionary<string, EndesaEntity.facturacion.AgoraManual> agoraManual;

            bool tiene_complemento_a01 = false;            
            bool sacar_portugal = true;

            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_totales;
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_agora;
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_Noagora;
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_agora_tam;
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_Noagora_tam;

           

            utilidades.Fechas utilfecha = new Fechas();

            DateTime fd = new DateTime();
            DateTime fd_tam = new DateTime();            

            try
            {
                // Ruta de la plantilla

                ss_pp.Update_Fecha_Inicio("Facturación", "PendML", "Agrupado Totales");

                FileInfo plantillaExcel = new FileInfo(System.Environment.CurrentDirectory + local_param.GetValue("plantilla_PS_AT_TAM"));

                sof.Contruye_Sofisticados();
                agoraManual = CargaAgoraManual(DateTime.Now, DateTime.Now);
                agoraPortugal = new contratacion.Agora_Portugal();

                FileInfo fileInfo = new FileInfo(ruta_salida_archivo);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);

                // #region Resumen

                var workSheet = excelPackage.Workbook.Worksheets["Resumen ES"];
                var headerCells = workSheet.Cells[1, 1, 1, 17];
                var headerFont = headerCells.Style.Font;

                fd = utilfecha.UltimoDiaHabilAnterior(
                utilfecha.UltimoDiaHabilAnterior(
                    utilfecha.UltimoDiaHabilAnterior(
                        utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()))));

                fd_tam = utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil());

                List<int> lista_empresas_ES = new List<int>();
                lista_empresas_ES.Add(20);
                lista_empresas_ES.Add(70);

                dic_totales = CargaTotales(lista_empresas_ES);
                dic_agora = CargaAgora_TAM(fd, lista_empresas_ES);
                dic_Noagora = CargaNoAgora_TAM(fd, lista_empresas_ES);

                dic_agora_tam = CargaAgora_TAM(fd_tam, lista_empresas_ES);
                dic_Noagora_tam = CargaNoAgora_TAM(fd_tam, lista_empresas_ES);
                

                bool firstOnly = true;
                int totales_dia = 0;
                double totales_dia_tam = 0;
                
                int totales_dia_facturacion = 0;
                double totales_dia_facturacion_tam = 0;
                int totales_dia_medida = 0;
                double totales_dia_medida_tam = 0;
                int[] total_general = new int[5];
                double[] total_general_tam = new double[2];
                int dif = 0;
                double dif_tam = 0;
                int dia = 0;
                int dia_tam = 0;
                

                c = 3;

                firstOnly = true;
                foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> p in dic_Noagora)
                {
                    totales_dia = 0;
                    totales_dia_medida = 0;
                    totales_dia_facturacion = 0;
                    

                    for (int i = 0; i < p.Value.Count; i++)                    
                        totales_dia = totales_dia + p.Value[i].num_cups;


                    total_general[dia] = totales_dia;

                    dia++;                    

                    #region diferencia
                    
                    //if (firstOnly)
                    //{
                    //    firstOnly = false;
                    //    f = 5;
                    //    c = 9;

                        

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.A. Endesa - Telemedida");

                        

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                        

                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.B. Endesa - TPL");


                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");


                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.E. NoEndesa - TPL");


                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");

                        

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                      

                    //    f++;

                       
                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", "03.A. CC Completa en el CS");

                       

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                        
                    //    f++;
                                              

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");

                        

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                        
                    //    f++;

                       

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");

                        

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                       
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");

                       

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                        

                    //    f++;

                       

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.A. No Validada");

                        

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                        

                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.B. Validada con Anomalías");

                       
                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                        

                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.C. Validada sin Anomalías");

                       

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.E. Modificada");

                        
                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                       

                    //    f++;


                    //    totales_dia_medida = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05.CC Rechazada por SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05. CC Enviada a SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", null);

                      



                    //    workSheet.Cells[f, c].Value = totales_dia_medida;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if (totales_dia_medida > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                       


                    //    f++;


                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");

                        

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                        
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(),  "08. El Punto no está Extraído", "08.B. Bloqueado");

                       

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                       
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.C. Otros");

                       

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                        

                    //    f++;
                        
                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", "09.A. El Punto está Extraído");

                        
                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                        
                    //    f++;                        

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", "10.A. Prefactura pendiente");

                       

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";                        
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                       
                    //    f++;


                    //    totales_dia_facturacion = 
                    //        Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", null)
                    //        + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", null)
                    //        + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", null);

                       
                    //    workSheet.Cells[f, c].Value = totales_dia_facturacion;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Font.Bold = true;
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if (totales_dia_facturacion > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                       
                    //    f++;

                    //    // Total No Ágora
                    //    workSheet.Cells[f, c].Value = totales_dia_facturacion + totales_dia_medida;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Font.Bold = true;
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;


                        
                    //    if ((totales_dia_facturacion + totales_dia_medida) > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                       

                    //    c = 3;

                    //}

                    #endregion


                    f = 4;
                    c++;

                    workSheet.Cells[f, c].Value = p.Key;
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", "01.B. Endesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                                                            
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;                                                          
                    
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                                                                               
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "05. CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;                    
                    
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", "07.A. No Validada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", "07.B. Validada con Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", "07.C. Validada sin Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", "07.E. Modificada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                    totales_dia_medida =  
                          Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "02. CC Rechazada por CS", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "03. CC Completa en CS", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "04. CC Enviada a SCE ML", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "05. CC Rechazada por SCE ML", null)                        
                        + Total_Pendiente(dic_Noagora, p.Key, "05. CC Enviada a SCE ML", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "06. CC Incompleta  SCE ML", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", null);


                    workSheet.Cells[f, c].Value = totales_dia_medida;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "08. El Punto no está Extraído", "08.B. Bloqueado");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "08. El Punto no está Extraído", "08.C. Otros");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;                                       
                    
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                    totales_dia_facturacion = Total_Pendiente(dic_Noagora, p.Key, "08. El Punto no está Extraído", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "09. El Punto está Extraído", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "10. Prefactura pendiente", null);

                    workSheet.Cells[f, c].Value = totales_dia_facturacion;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = totales_dia;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;

                    
                }

                c = 8;
                firstOnly = true;
                foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> p in dic_Noagora_tam)
                {
                    totales_dia_tam = 0;
                    totales_dia_medida_tam = 0;
                    totales_dia_facturacion_tam = 0;


                    for (int i = 0; i < p.Value.Count; i++)
                        totales_dia_tam = totales_dia_tam + p.Value[i].tam;


                    total_general_tam[dia_tam] = totales_dia_tam;

                    dia_tam++;

                    

                    f = 4;
                    c++;

                    workSheet.Cells[f, c].Value = p.Key;
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "01. Pendiente de medida", "01.B. Endesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "07. LTP SCE", "07.A. No Validada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "07. LTP SCE", "07.B. Validada con Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "07. LTP SCE", "07.C. Validada sin Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "07. LTP SCE", "07.E. Modificada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;


                    totales_dia_medida_tam = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "01. Pendiente de medida", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "02. CC Rechazada por CS", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "03. CC Completa en CS", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "04. CC Enviada a SCE ML", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "05.CC Rechazada por SCE ML", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "05. CC Enviada a SCE ML", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "06. CC Incompleta  SCE ML", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "07. LTP SCE", null);


                    workSheet.Cells[f, c].Value = totales_dia_medida_tam;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "08. El Punto no está Extraído", "08.B. Bloqueado");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "08. El Punto no está Extraído", "08.C. Otros");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    totales_dia_facturacion_tam = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "08. El Punto no está Extraído", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "09. El Punto está Extraído", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "10. Prefactura pendiente", null);

                    workSheet.Cells[f, c].Value = totales_dia_facturacion_tam;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = totales_dia_tam;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;

                }


                headerCells = workSheet.Cells[1, 1, 1, 30];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;
                var allCells = workSheet.Cells[1, 1, 50, 50];
                

                c = 3;
                dia = 0;
                firstOnly = true;
                foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> p in dic_agora)
                {
                    totales_dia = 0;
                    totales_dia_medida = 0;
                    totales_dia_facturacion = 0;

                    for (int i = 0; i < p.Value.Count; i++)
                        totales_dia = totales_dia + p.Value[i].num_cups;

                    total_general[dia] = total_general[dia] + totales_dia;
                    dia++;

                    #region Evolucion

                    //if (firstOnly)
                    //{
                    //    firstOnly = false;
                    //    f = 26;
                    //    c = 9;

                    //    //total_medida_no_agora = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabil(), 
                    //    //    utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()), "01. Pendiente de medida", null);

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.A. Endesa - Telemedida");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.B. Endesa - TPL");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.E. NoEndesa - TPL");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    //total_medida_no_agora += Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabil(), 
                    //    //    utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()), "02. CC Rechazada por CS", null);


                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    //total_medida_no_agora += Total_Pendiente_DIFF(dic_Noagora, 
                    //    //    utilfecha.UltimoDiaHabil(), utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()), "03. CC Completa en CS", null);


                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", "03.A. CC Completa en el CS");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    //total_medida_no_agora += Total_Pendiente_DIFF(dic_totales, utilfecha.UltimoDiaHabil(), 
                    //    //    utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()), "04. CC Enviada a SCE ML", null);

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    //total_medida_no_agora += Total_Pendiente_DIFF(dic_totales, utilfecha.UltimoDiaHabil(), 
                    //    //    utilfecha.UltimoDiaHabilAnterior(p.Key), "06. CC Incompleta  SCE ML", null);


                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    //total_medida_no_agora += Total_Pendiente_DIFF(dic_totales, utilfecha.UltimoDiaHabil(), 
                    //    //    utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()), "07. LTP SCE", null);


                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.A. No Validada");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.B. Validada con Anomalías");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.C. Validada sin Anomalías");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.E. Modificada");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    totales_dia_medida = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05.CC Rechazada por SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05. CC Enviada a SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", null);



                    //    workSheet.Cells[f, c].Value = totales_dia_medida;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, 12].Style.Font.Bold = true;
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;



                    //    //total_facturacion_no_agora = Total_Pendiente_DIFF(dic_totales, utilfecha.UltimoDiaHabil(), 
                    //    //    utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()), "08. El Punto no está Extraído", null);

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.B. Bloqueado");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.C. Otros");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                        
                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", "09.A. El Punto está Extraído");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                        
                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", "10.A. Prefactura pendiente");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    totales_dia_facturacion =
                    //        Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", null)
                    //        + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", null)
                    //        + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", null);

                    //    workSheet.Cells[f, c].Value = totales_dia_facturacion;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Font.Bold = true;
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    // Total No Ágora
                    //    workSheet.Cells[f, c].Value = totales_dia_facturacion + totales_dia_medida;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Font.Bold = true;
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    c = 3;



                    //}

                    #endregion Evolucion
                    f = 26;
                    c++;


                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", "01.B. Endesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", "07.A. No Validada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", "07.B. Validada con Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", "07.C. Validada sin Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", "07.E. Modificada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                    totales_dia_medida = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", null)
                        + Total_Pendiente(dic_agora, p.Key, "02. CC Rechazada por CS", null)
                        + Total_Pendiente(dic_agora, p.Key, "03. CC Completa en CS", null)
                        + Total_Pendiente(dic_agora, p.Key, "04. CC Enviada a SCE ML", null)
                        + Total_Pendiente(dic_agora, p.Key, "05.CC Rechazada por SCE ML", null)
                        + Total_Pendiente(dic_agora, p.Key, "05. CC Enviada a SCE ML", null)
                        + Total_Pendiente(dic_agora, p.Key, "06. CC Incompleta  SCE ML", null)
                        + Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", null);


                    workSheet.Cells[f, c].Value = totales_dia_medida;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "08. El Punto no está Extraído", "08.B. Bloqueado");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "08. El Punto no está Extraído", "08.C. Otros");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                    totales_dia_facturacion = Total_Pendiente(dic_agora, p.Key, "08. El Punto no está Extraído", null)
                        + Total_Pendiente(dic_agora, p.Key, "09. El Punto está Extraído", null)
                        + Total_Pendiente(dic_agora, p.Key, "10. Prefactura pendiente", null);

                    workSheet.Cells[f, c].Value = totales_dia_facturacion;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = totales_dia;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;




                }

                c = 8;
                firstOnly = true;
                dia_tam = 0;
                foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> p in dic_agora_tam)
                {
                    totales_dia_tam = 0;
                    totales_dia_medida_tam = 0;
                    totales_dia_facturacion_tam = 0;


                    for (int i = 0; i < p.Value.Count; i++)
                        totales_dia_tam = totales_dia_tam + p.Value[i].tam;


                    total_general_tam[dia_tam] = total_general_tam[dia_tam] + totales_dia_tam;

                    dia_tam++;

                    #region diferencia

                    //if (firstOnly)
                    //{
                    //    firstOnly = false;
                    //    f = 26;
                    //    c = 9;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //       utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.A. Endesa - Telemedida");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.B. Endesa - TPL");



                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.E. NoEndesa - TPL");



                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;




                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", "03.A. CC Completa en el CS");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");



                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");



                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.A. No Validada");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.B. Validada con Anomalías");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.C. Validada sin Anomalías");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.E. Modificada");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;



                    //    totales_dia_medida_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05.CC Rechazada por SCE ML", null)                        
                    //    + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", null);




                    //    workSheet.Cells[f, 12].Value = totales_dia_medida_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Font.Bold = true;
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (totales_dia_medida_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //       utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.B. Bloqueado");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.C. Otros");

                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;

                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", "09.A. El Punto está Extraído");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", "10.A. Prefactura pendiente");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;



                    //    totales_dia_facturacion_tam =
                    //        Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", null)
                    //        + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", null)
                    //        + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", null);



                    //    workSheet.Cells[f, 12].Value = totales_dia_facturacion_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Font.Bold = true;
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if (totales_dia_facturacion_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    // Total No Ágora TAM
                    //    workSheet.Cells[f, 12].Value = totales_dia_facturacion_tam + totales_dia_medida_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Font.Bold = true;
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if ((totales_dia_facturacion_tam + totales_dia_medida_tam) > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    c = 9;

                    //}

                    #endregion

                    f = 25;
                    c++;
                    
                    f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "01. Pendiente de medida", "01.B. Endesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "07. LTP SCE", "07.A. No Validada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "07. LTP SCE", "07.B. Validada con Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "07. LTP SCE", "07.C. Validada sin Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "07. LTP SCE", "07.E. Modificada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;


                    totales_dia_medida_tam = Total_Pendiente_TAM(dic_agora_tam, p.Key, "01. Pendiente de medida", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "02. CC Rechazada por CS", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "03. CC Completa en CS", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "04. CC Enviada a SCE ML", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "05. CC Rechazada por SCE ML", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "05. CC Enviada a SCE ML", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "06. CC Incompleta  SCE ML", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "07. LTP SCE", null);


                    workSheet.Cells[f, c].Value = totales_dia_medida_tam;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "08. El Punto no está Extraído", "08.B. Bloqueado");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "08. El Punto no está Extraído", "08.C. Otros");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    totales_dia_facturacion_tam = Total_Pendiente_TAM(dic_agora_tam, p.Key, "08. El Punto no está Extraído", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "09. El Punto está Extraído", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "10. Prefactura pendiente", null);

                    workSheet.Cells[f, c].Value = totales_dia_facturacion_tam;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = totales_dia_tam;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;


                }

                c = 3;
                for(int j = 0; j < 5; j++)
                {
                    c++;
                    workSheet.Cells[47, c].Value = total_general[j];
                    workSheet.Cells[47, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[47, c].Style.Font.Bold = true;
                    
                }

                c = 8;
                for (int j = 0; j < 2; j++)
                {
                    c++;
                    workSheet.Cells[47, c].Value = total_general_tam[j];
                    workSheet.Cells[47, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[47, c].Style.Font.Bold = true;

                }



                //c = 9;
                //int total_evolucion = total_general[4] - total_general[3];
                //workSheet.Cells[47, c].Value = total_evolucion;
                //workSheet.Cells[47, c].Style.Numberformat.Format = "#,##0";
                //workSheet.Cells[47, c].Style.Font.Bold = true;
                //workSheet.Cells[47, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //if(total_evolucion > 0)
                //    workSheet.Cells[47, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                //else
                //    workSheet.Cells[47, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                //c = 12;
                //double total_evolucion_tam = total_general_tam[1] - total_general_tam[0];
                //workSheet.Cells[47, c].Value = total_evolucion_tam;
                //workSheet.Cells[47, c].Style.Numberformat.Format = "#,##0.00";
                //workSheet.Cells[47, c].Style.Font.Bold = true;
                //workSheet.Cells[47, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //if (total_evolucion_tam > 0)
                //    workSheet.Cells[47, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                //else
                //    workSheet.Cells[47, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));



                #region Detalle ES

                f = 1;
                c = 1;

                workSheet = excelPackage.Workbook.Worksheets["Detalle ES"];
                headerCells = workSheet.Cells[1, 1, 1, 17];
                headerFont = headerCells.Style.Font;

                workSheet.View.FreezePanes(2, 1);

                headerFont.Bold = true;
                workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                workSheet.Cells[f, c].Value = "NIF"; c++;
                workSheet.Cells[f, c].Value = "CLIENTE"; c++;
                workSheet.Cells[f, c].Value = "FALTACONT"; c++;
                workSheet.Cells[f, c].Value = "FPSERCON"; c++;
                workSheet.Cells[f, c].Value = "CUPS13"; c++;
                workSheet.Cells[f, c].Value = "CUPS20"; c++;
                workSheet.Cells[f, c].Value = "TARIFA"; c++;
                workSheet.Cells[f, c].Value = "CONTRATO"; c++;
                workSheet.Cells[f, c].Value = "MES"; c++;
                workSheet.Cells[f, c].Value = "DISTRIBUIDORA"; c++;
                workSheet.Cells[f, c].Value = "ESTADO"; c++;
                workSheet.Cells[f, c].Value = "SUBESTADO"; c++;
                workSheet.Cells[f, c].Value = "MULTIMEDIDA"; c++;
                workSheet.Cells[f, c].Value = "TAM"; c++;
                workSheet.Cells[f, c].Value = "MESES PDTES FACTURAR"; c++;
                workSheet.Cells[f, c].Value = "IMPORTE PDTE FACTURAR"; c++;
                workSheet.Cells[f, c].Value = "ÁGORA";


                for (int i = 1; i <= c; i++)
                {
                    workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }


                strSql = "SELECT empresa_titular, ps.NIF, ps.Cliente, ps.fAltaCont, ps.FPSERCON,"
                    + " p.cups13, ps.TARIFA, substr(ps.CUPS22, 1, 20) as CUPS20,"
                    + " contrato, mes, distribuidora, estado, subestado, multimedida, if (t.TAM IS NULL, 0, t.tam) AS TAM"
                    + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar_agrupado p"
                    + " LEFT OUTER JOIN cont.PS_AT ps ON"
                    + " ps.IDU = p.cups13"
                    + " LEFT OUTER JOIN fact.tam t ON"
                    + " t.CEMPTITU = p.empresa_titular AND"
                    + " t.CNIFDNIC = ps.NIF AND"
                    + " t.CUPS13 = ps.IDU"
                    + " WHERE p.empresa_titular <> 80"
                    + " and p.cups13 <> 'VZZ0605131910'"
                    + " GROUP BY p.cups13";


                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    f++;
                    c = 1;
                    if (r["empresa_titular"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["empresa_titular"]);
                    c++;
                    if (r["NIF"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["NIF"].ToString();
                    c++;
                    if (r["Cliente"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["Cliente"].ToString();
                    c++;

                    if (r["fAltaCont"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["fAltaCont"]);
                    }

                    c++;
                    if (r["FPSERCON"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["FPSERCON"]);

                    c++;
                    if (r["cups13"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["cups13"].ToString();

                    c++;
                    if (r["CUPS20"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CUPS20"].ToString();

                    c++;
                    if (r["TARIFA"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                    c++;
                    if (r["contrato"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt64(r["contrato"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = "###0";
                    }

                    c++;
                    if (r["mes"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["mes"]);
                        aniomes = Convert.ToInt32(r["mes"]);
                        fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                            Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                    }

                    c++;
                    if (r["distribuidora"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["distribuidora"].ToString();
                    c++;
                    if (r["estado"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["estado"].ToString();
                    c++;
                    if (r["subestado"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["subestado"].ToString();
                    c++;
                    if (r["multimedida"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["multimedida"].ToString();
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    c++;

                    if (r["TAM"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    }

                    c++;

                    if (r["mes"] != System.DBNull.Value)
                    {
                        //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                        meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                        workSheet.Cells[f, c].Value = meses_pdtes; c++;
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]) * meses_pdtes;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    else
                    {
                        c++; c++;
                    }

                    // AGORA
                    if (r["cups13"] != System.DBNull.Value)
                    {
                        tiene_complemento_a01 = sof.EsSofisticado(r["cups13"].ToString());
                        if (!tiene_complemento_a01)
                        {
                            if (r["CUPS20"] != System.DBNull.Value)
                            {
                                EndesaEntity.facturacion.AgoraManual o;
                                if (agoraManual.TryGetValue(r["CUPS20"].ToString(), out o))
                                    tiene_complemento_a01 = true;
                            }

                        }

                        if (tiene_complemento_a01 && Convert.ToInt32(r["empresa_titular"]) != 70)
                        {
                            workSheet.Cells[f, c].Value = "Sí";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else
                        {
                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    }





                }
                db.CloseConnection();

                headerCells = workSheet.Cells[1, 1, 1, c];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;
                allCells = workSheet.Cells[1, 1, f, c];

                allCells.AutoFitColumns();

                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:R1"].AutoFilter = true;
                allCells.AutoFitColumns();

                #endregion

                #region Resumen Portugal MT
                List<int> lista_empresas_PT = new List<int>();
                lista_empresas_PT.Add(80);                

                dic_totales = CargaTotales(lista_empresas_PT);
                dic_agora = CargaAgora_TAM(fd, lista_empresas_PT);
                dic_Noagora = CargaNoAgora_TAM(fd, lista_empresas_PT);

                dic_agora_tam = CargaAgora_TAM(fd_tam, lista_empresas_PT);
                dic_Noagora_tam = CargaNoAgora_TAM(fd_tam, lista_empresas_PT);

                workSheet = excelPackage.Workbook.Worksheets["Resumen Portugal MT"];

                firstOnly = true;
                totales_dia = 0;
                totales_dia_tam = 0;

                totales_dia_facturacion = 0;
                totales_dia_facturacion_tam = 0;
                totales_dia_medida = 0;
                totales_dia_medida_tam = 0;
                total_general = new int[5];
                total_general_tam = new double[2];
                dif = 0;
                dif_tam = 0;
                dia = 0;
                dia_tam = 0;


                c = 3;

                firstOnly = true;
                foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> p in dic_Noagora)
                {
                    totales_dia = 0;
                    totales_dia_medida = 0;
                    totales_dia_facturacion = 0;


                    for (int i = 0; i < p.Value.Count; i++)
                        totales_dia = totales_dia + p.Value[i].num_cups;


                    total_general[dia] = totales_dia;

                    dia++;

                    #region diferencia

                    //if (firstOnly)
                    //{
                    //    firstOnly = false;
                    //    f = 5;
                    //    c = 9;



                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.A. Endesa - Telemedida");



                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));




                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.B. Endesa - TPL");


                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");


                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.E. NoEndesa - TPL");


                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");



                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));



                    //    f++;


                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", "03.A. CC Completa en el CS");



                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;


                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");



                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;



                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");



                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");



                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));



                    //    f++;



                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.A. No Validada");



                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));



                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.B. Validada con Anomalías");


                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));



                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.C. Validada sin Anomalías");



                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.E. Modificada");


                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));



                    //    f++;


                    //    totales_dia_medida = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05.CC Rechazada por SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05. CC Enviada a SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", null);





                    //    workSheet.Cells[f, c].Value = totales_dia_medida;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if (totales_dia_medida > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));




                    //    f++;


                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");



                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(),  "08. El Punto no está Extraído", "08.B. Bloqueado");



                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.C. Otros");



                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));



                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", "09.A. El Punto está Extraído");


                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;                        

                    //    dif = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", "10.A. Prefactura pendiente");



                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";                        
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;


                    //    totales_dia_facturacion = 
                    //        Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", null)
                    //        + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", null)
                    //        + Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", null);


                    //    workSheet.Cells[f, c].Value = totales_dia_facturacion;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Font.Bold = true;
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if (totales_dia_facturacion > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));



                    //    f++;

                    //    // Total No Ágora
                    //    workSheet.Cells[f, c].Value = totales_dia_facturacion + totales_dia_medida;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Font.Bold = true;
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;



                    //    if ((totales_dia_facturacion + totales_dia_medida) > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));



                    //    c = 3;

                    //}

                    #endregion


                    f = 4;
                    c++;

                    workSheet.Cells[f, c].Value = p.Key;
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", "01.B. Endesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", "07.A. No Validada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", "07.B. Validada con Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", "07.C. Validada sin Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", "07.E. Modificada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                    totales_dia_medida = Total_Pendiente(dic_Noagora, p.Key, "01. Pendiente de medida", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "02. CC Rechazada por CS", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "03. CC Completa en CS", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "04. CC Enviada a SCE ML", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "05.CC Rechazada por SCE ML", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "05. CC Enviada a SCE ML", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "06. CC Incompleta  SCE ML", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "07. LTP SCE", null);


                    workSheet.Cells[f, c].Value = totales_dia_medida;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "08. El Punto no está Extraído", "08.B. Bloqueado");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "08. El Punto no está Extraído", "08.C. Otros");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_Noagora, p.Key, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                    totales_dia_facturacion = Total_Pendiente(dic_Noagora, p.Key, "08. El Punto no está Extraído", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "09. El Punto está Extraído", null)
                        + Total_Pendiente(dic_Noagora, p.Key, "10. Prefactura pendiente", null);

                    workSheet.Cells[f, c].Value = totales_dia_facturacion;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = totales_dia;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;


                }

                c = 8;
                firstOnly = true;
                foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> p in dic_Noagora_tam)
                {
                    totales_dia_tam = 0;
                    totales_dia_medida_tam = 0;
                    totales_dia_facturacion_tam = 0;


                    for (int i = 0; i < p.Value.Count; i++)
                        totales_dia_tam = totales_dia_tam + p.Value[i].tam;


                    total_general_tam[dia_tam] = totales_dia_tam;

                    dia_tam++;

                    #region diferencia

                    //if (firstOnly)
                    //{
                    //    firstOnly = false;
                    //    f = 5;
                    //    c = 9;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //       utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.A. Endesa - Telemedida");                       


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.B. Endesa - TPL");



                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.E. NoEndesa - TPL");



                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;




                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", "03.A. CC Completa en el CS");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");



                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");



                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.A. No Validada");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.B. Validada con Anomalías");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.C. Validada sin Anomalías");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.E. Modificada");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;



                    //    totales_dia_medida_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05. CC Rechazada por SCE ML", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05. CC Enviada a SCE ML", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", null);

                    //    workSheet.Cells[f, 12].Value = totales_dia_medida_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Font.Bold = true;
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (totales_dia_medida_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //       utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.B. Bloqueado");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.C. Otros");

                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;

                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", "09.A. El Punto está Extraído");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", "10.A. Prefactura pendiente");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;



                    //    totales_dia_facturacion_tam =
                    //        Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", null)
                    //        + Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", null)
                    //        + Total_Pendiente_DIFF_TAM(dic_Noagora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", null);



                    //    workSheet.Cells[f, 12].Value = totales_dia_facturacion_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Font.Bold = true;
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if (totales_dia_facturacion_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    // Total No Ágora TAM
                    //    workSheet.Cells[f, 12].Value = totales_dia_facturacion_tam + totales_dia_medida_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Font.Bold = true;
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if ((totales_dia_facturacion_tam + totales_dia_medida_tam) > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    c = 9;

                    //}

                    #endregion

                    f = 4;
                    c++;

                    workSheet.Cells[f, c].Value = p.Key;
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "01. Pendiente de medida", "01.B. Endesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "07. LTP SCE", "07.A. No Validada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "07. LTP SCE", "07.B. Validada con Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "07. LTP SCE", "07.C. Validada sin Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "07. LTP SCE", "07.E. Modificada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;


                    totales_dia_medida_tam = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "01. Pendiente de medida", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "02. CC Rechazada por CS", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "03. CC Completa en CS", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "04. CC Enviada a SCE ML", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "05.CC Rechazada por SCE ML", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "05. CC Enviada a SCE ML", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "06. CC Incompleta  SCE ML", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "07. LTP SCE", null);


                    workSheet.Cells[f, c].Value = totales_dia_medida_tam;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "08. El Punto no está Extraído", "08.B. Bloqueado");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "08. El Punto no está Extraído", "08.C. Otros");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    totales_dia_facturacion_tam = Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "08. El Punto no está Extraído", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "09. El Punto está Extraído", null)
                        + Total_Pendiente_TAM(dic_Noagora_tam, p.Key, "10. Prefactura pendiente", null);

                    workSheet.Cells[f, c].Value = totales_dia_facturacion_tam;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = totales_dia_tam;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;

                }


                headerCells = workSheet.Cells[1, 1, 1, 30];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;
                allCells = workSheet.Cells[1, 1, 50, 50];


                c = 3;
                dia = 0;
                firstOnly = true;
                foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> p in dic_agora)
                {
                    totales_dia = 0;
                    totales_dia_medida = 0;
                    totales_dia_facturacion = 0;

                    for (int i = 0; i < p.Value.Count; i++)
                        totales_dia = totales_dia + p.Value[i].num_cups;

                    total_general[dia] = total_general[dia] + totales_dia;
                    dia++;

                    #region Evolucion

                    //if (firstOnly)
                    //{
                    //    firstOnly = false;
                    //    f = 26;
                    //    c = 9;

                    //    //total_medida_no_agora = Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabil(), 
                    //    //    utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()), "01. Pendiente de medida", null);

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.A. Endesa - Telemedida");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.B. Endesa - TPL");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.E. NoEndesa - TPL");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    //total_medida_no_agora += Total_Pendiente_DIFF(dic_Noagora, utilfecha.UltimoDiaHabil(), 
                    //    //    utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()), "02. CC Rechazada por CS", null);


                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    //total_medida_no_agora += Total_Pendiente_DIFF(dic_Noagora, 
                    //    //    utilfecha.UltimoDiaHabil(), utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()), "03. CC Completa en CS", null);


                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", "03.A. CC Completa en el CS");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    //total_medida_no_agora += Total_Pendiente_DIFF(dic_totales, utilfecha.UltimoDiaHabil(), 
                    //    //    utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()), "04. CC Enviada a SCE ML", null);

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    //total_medida_no_agora += Total_Pendiente_DIFF(dic_totales, utilfecha.UltimoDiaHabil(), 
                    //    //    utilfecha.UltimoDiaHabilAnterior(p.Key), "06. CC Incompleta  SCE ML", null);


                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    //total_medida_no_agora += Total_Pendiente_DIFF(dic_totales, utilfecha.UltimoDiaHabil(), 
                    //    //    utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()), "07. LTP SCE", null);


                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.A. No Validada");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.B. Validada con Anomalías");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.C. Validada sin Anomalías");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.E. Modificada");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    totales_dia_medida = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05.CC Rechazada por SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05. CC Enviada a SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", null)
                    //    + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", null);



                    //    workSheet.Cells[f, c].Value = totales_dia_medida;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, 12].Style.Font.Bold = true;
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;



                    //    //total_facturacion_no_agora = Total_Pendiente_DIFF(dic_totales, utilfecha.UltimoDiaHabil(), 
                    //    //    utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()), "08. El Punto no está Extraído", null);

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.B. Bloqueado");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.C. Otros");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;



                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", "09.A. El Punto está Extraído");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    dif = Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", "10.A. Prefactura pendiente");

                    //    workSheet.Cells[f, c].Value = dif;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    totales_dia_facturacion =
                    //        Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", null)
                    //        + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", null)
                    //        + Total_Pendiente_DIFF(dic_agora, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", null);

                    //    workSheet.Cells[f, c].Value = totales_dia_facturacion;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Font.Bold = true;
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    // Total No Ágora
                    //    workSheet.Cells[f, c].Value = totales_dia_facturacion + totales_dia_medida;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //    workSheet.Cells[f, c].Style.Font.Bold = true;
                    //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif > 0)
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    c = 3;



                    //}

                    #endregion Evolucion
                    f = 26;
                    c++;


                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", "01.B. Endesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", "07.A. No Validada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", "07.B. Validada con Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", "07.C. Validada sin Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", "07.E. Modificada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                    totales_dia_medida = Total_Pendiente(dic_agora, p.Key, "01. Pendiente de medida", null)
                        + Total_Pendiente(dic_agora, p.Key, "02. CC Rechazada por CS", null)
                        + Total_Pendiente(dic_agora, p.Key, "03. CC Completa en CS", null)
                        + Total_Pendiente(dic_agora, p.Key, "04. CC Enviada a SCE ML", null)
                        + Total_Pendiente(dic_agora, p.Key, "05.CC Rechazada por SCE ML", null)
                        + Total_Pendiente(dic_agora, p.Key, "05. CC Enviada a SCE ML", null)
                        + Total_Pendiente(dic_agora, p.Key, "06. CC Incompleta  SCE ML", null)
                        + Total_Pendiente(dic_agora, p.Key, "07. LTP SCE", null);


                    workSheet.Cells[f, c].Value = totales_dia_medida;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "08. El Punto no está Extraído", "08.B. Bloqueado");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "08. El Punto no está Extraído", "08.C. Otros");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(dic_agora, p.Key, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                    totales_dia_facturacion = Total_Pendiente(dic_agora, p.Key, "08. El Punto no está Extraído", null)
                        + Total_Pendiente(dic_agora, p.Key, "09. El Punto está Extraído", null)
                        + Total_Pendiente(dic_agora, p.Key, "10. Prefactura pendiente", null);

                    workSheet.Cells[f, c].Value = totales_dia_facturacion;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = totales_dia;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;




                }

                c = 8;
                firstOnly = true;
                dia_tam = 0;
                foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> p in dic_agora_tam)
                {
                    totales_dia_tam = 0;
                    totales_dia_medida_tam = 0;
                    totales_dia_facturacion_tam = 0;


                    for (int i = 0; i < p.Value.Count; i++)
                        totales_dia_tam = totales_dia_tam + p.Value[i].tam;


                    total_general_tam[dia_tam] = total_general_tam[dia_tam] + totales_dia_tam;

                    dia_tam++;

                    #region diferencia

                    //if (firstOnly)
                    //{
                    //    firstOnly = false;
                    //    f = 26;
                    //    c = 9;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //       utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.A. Endesa - Telemedida");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.B. Endesa - TPL");



                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", "01.E. NoEndesa - TPL");



                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;




                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", "03.A. CC Completa en el CS");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");



                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");



                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.A. No Validada");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.B. Validada con Anomalías");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.C. Validada sin Anomalías");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", "07.E. Modificada");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;



                    //    totales_dia_medida_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "01. Pendiente de medida", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "02. CC Rechazada por CS", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "03. CC Completa en CS", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "04. CC Enviada a SCE ML", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "05.CC Rechazada por SCE ML", null)                        
                    //    + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "06. CC Incompleta  SCE ML", null)
                    //    + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "07. LTP SCE", null);




                    //    workSheet.Cells[f, 12].Value = totales_dia_medida_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Font.Bold = true;
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (totales_dia_medida_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));


                    //    f++;



                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //       utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.B. Bloqueado");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", "08.C. Otros");

                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;

                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", "09.A. El Punto está Extraído");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;


                    //    dif_tam = Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", "10.A. Prefactura pendiente");


                    //    workSheet.Cells[f, 12].Value = dif_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    if (dif_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;



                    //    totales_dia_facturacion_tam =
                    //        Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "08. El Punto no está Extraído", null)
                    //        + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "09. El Punto está Extraído", null)
                    //        + Total_Pendiente_DIFF_TAM(dic_agora_tam, utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()),
                    //        utilfecha.UltimoDiaHabil(), "10. Prefactura pendiente", null);



                    //    workSheet.Cells[f, 12].Value = totales_dia_facturacion_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Font.Bold = true;
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if (totales_dia_facturacion_tam > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));

                    //    f++;


                    //    // Total No Ágora TAM
                    //    workSheet.Cells[f, 12].Value = totales_dia_facturacion_tam + totales_dia_medida_tam;
                    //    workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                    //    workSheet.Cells[f, 12].Style.Font.Bold = true;
                    //    workSheet.Cells[f, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                    //    if ((totales_dia_facturacion_tam + totales_dia_medida_tam) > 0)
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(251, 21, 21));
                    //    else
                    //        workSheet.Cells[f, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(35, 146, 14));
                    //    f++;

                    //    c = 9;

                    //}

                    #endregion

                    f = 25;
                    c++;

                    f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "01. Pendiente de medida", "01.B. Endesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "07. LTP SCE", "07.A. No Validada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "07. LTP SCE", "07.B. Validada con Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "07. LTP SCE", "07.C. Validada sin Anomalías");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "07. LTP SCE", "07.E. Modificada");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;


                    totales_dia_medida_tam = Total_Pendiente_TAM(dic_agora_tam, p.Key, "01. Pendiente de medida", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "02. CC Rechazada por CS", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "03. CC Completa en CS", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "04. CC Enviada a SCE ML", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "05. CC Rechazada por SCE ML", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "05. CC Enviada a SCE ML", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "06. CC Incompleta  SCE ML", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "07. LTP SCE", null);


                    workSheet.Cells[f, c].Value = totales_dia_medida_tam;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "08. El Punto no está Extraído", "08.B. Bloqueado");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "08. El Punto no está Extraído", "08.C. Otros");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(dic_agora_tam, p.Key, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    totales_dia_facturacion_tam = Total_Pendiente_TAM(dic_agora_tam, p.Key, "08. El Punto no está Extraído", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "09. El Punto está Extraído", null)
                        + Total_Pendiente_TAM(dic_agora_tam, p.Key, "10. Prefactura pendiente", null);

                    workSheet.Cells[f, c].Value = totales_dia_facturacion_tam;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = totales_dia_tam;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;


                }

                c = 3;
                for (int j = 0; j < 5; j++)
                {
                    c++;
                    workSheet.Cells[47, c].Value = total_general[j];
                    workSheet.Cells[47, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[47, c].Style.Font.Bold = true;

                }

                c = 8;
                for (int j = 0; j < 2; j++)
                {
                    c++;
                    workSheet.Cells[47, c].Value = total_general_tam[j];
                    workSheet.Cells[47, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[47, c].Style.Font.Bold = true;

                }




                #endregion

                #region 80_MT_Portugal
                if (sacar_portugal)
                {
                    c = 1;
                    f = 1;

                    workSheet = excelPackage.Workbook.Worksheets["Detalle Portugal MT"];
                    headerCells = workSheet.Cells[1, 1, 1, 17];
                    headerFont = headerCells.Style.Font;


                    workSheet.View.FreezePanes(2, 1);

                    headerFont.Bold = true;
                    workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                    workSheet.Cells[f, c].Value = "NIF"; c++;
                    workSheet.Cells[f, c].Value = "CLIENTE"; c++;
                    workSheet.Cells[f, c].Value = "FALTACONT"; c++;
                    workSheet.Cells[f, c].Value = "FPSERCON"; c++;
                    workSheet.Cells[f, c].Value = "CUPS13"; c++;
                    workSheet.Cells[f, c].Value = "CUPS20"; c++;
                    workSheet.Cells[f, c].Value = "TARIFA"; c++;
                    workSheet.Cells[f, c].Value = "CONTRATO"; c++;
                    workSheet.Cells[f, c].Value = "MES"; c++;
                    workSheet.Cells[f, c].Value = "DISTRIBUIDORA"; c++;
                    workSheet.Cells[f, c].Value = "ESTADO"; c++;
                    workSheet.Cells[f, c].Value = "SUBESTADO"; c++;
                    workSheet.Cells[f, c].Value = "MULTIMEDIDA"; c++;
                    workSheet.Cells[f, c].Value = "TAM"; c++;
                    workSheet.Cells[f, c].Value = "MESES PDTES FACTURAR"; c++;
                    workSheet.Cells[f, c].Value = "IMPORTE PDTE FACTURAR"; c++;
                    workSheet.Cells[f, c].Value = "ÁGORA";

                    for (int i = 1; i <= c; i++)
                    {
                        workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }


                    DateTime ultima_carga = new DateTime();
                    strSql = "SELECT MAX(f_ult_mod) f_ult_mod FROM contratos_ps_mt";
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        ultima_carga = Convert.ToDateTime(r["f_ult_mod"]);
                    }
                    db.CloseConnection();

                    strSql = "SELECT empresa_titular, ps.faltacon, ps.CNIFDNIC, ps.DAPERSOC,"
                        + " p.cups13, ps.CTARIFA as tarifa, ps.CUPSREE as cups20,"
                        + " contrato, mes, distribuidora, estado, subestado, multimedida, if (t.TAM IS NULL, 0, t.tam) AS TAM"
                        + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar_agrupado p"
                        + " LEFT OUTER JOIN cont.contratos_ps_mt ps ON"
                        + " ps.CUPS = p.cups13"
                        + " LEFT OUTER JOIN fact.tam t ON"
                        //+ " t.CEMPTITU = p.empresa_titular AND"
                        //+ " t.CNIFDNIC = ps.NIF AND"
                        + " t.CUPS13 = ps.CUPS"
                        + " WHERE p.empresa_titular = 80 and"
                        + " ps.f_ult_mod >= '" + ultima_carga.ToString("yyyy-MM-dd") + "'"
                        + " GROUP BY p.cups13";


                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        f++;
                        c = 1;
                        if (r["empresa_titular"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["empresa_titular"]);
                        c++;

                        if (r["CNIFDNIC"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CNIFDNIC"].ToString();
                        c++;

                        if (r["DAPERSOC"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["DAPERSOC"].ToString();
                        c++;

                        if (r["faltacon"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["faltacon"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }

                        c++;

                        //if (r["FPSERCON"] != System.DBNull.Value)
                        //    workSheet.Cells[f, c].Value = Convert.ToInt32(r["FPSERCON"]);

                        c++;
                        if (r["cups13"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cups13"].ToString();

                        c++;
                        if (r["cups20"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cups20"].ToString();

                        c++;
                        if (r["TARIFA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                        c++;
                        if (r["contrato"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt64(r["contrato"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "###0";
                        }

                        c++;
                        if (r["mes"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["mes"]);
                            aniomes = Convert.ToInt32(r["mes"]);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                        }

                        c++;
                        if (r["distribuidora"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["distribuidora"].ToString();
                        c++;
                        if (r["estado"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["estado"].ToString();
                        c++;
                        if (r["subestado"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["subestado"].ToString();
                        c++;
                        if (r["multimedida"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["multimedida"].ToString();
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        c++;
                        if (r["TAM"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        }


                        c++;

                        if (r["mes"] != System.DBNull.Value)
                        {
                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                            workSheet.Cells[f, c].Value = meses_pdtes; c++;
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]) * meses_pdtes;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        }
                        else
                        {
                            c++; c++;
                        }


                        // AGORA
                        if (r["cups13"] != System.DBNull.Value)
                        {
                            tiene_complemento_a01 = sof.EsSofisticado(r["cups13"].ToString());

                            if (!tiene_complemento_a01)
                            {
                                if (r["cups20"] != System.DBNull.Value)
                                {
                                    EndesaEntity.facturacion.AgoraManual o;
                                    if (agoraManual.TryGetValue(r["CUPS20"].ToString(), out o))
                                        tiene_complemento_a01 = true;
                                }

                                if (!tiene_complemento_a01)
                                    tiene_complemento_a01 = agoraPortugal.EsAgora(r["cups13"].ToString());

                            }

                            if (tiene_complemento_a01)
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            else
                            {
                                workSheet.Cells[f, c].Value = "No";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }

                        }


                    }
                    db.CloseConnection();

                    headerCells = workSheet.Cells[1, 1, 1, c];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;
                    allCells = workSheet.Cells[1, 1, f, c];

                    allCells.AutoFitColumns();

                    workSheet.View.FreezePanes(2, 1);
                    workSheet.Cells["A1:R1"].AutoFilter = true;
                    allCells.AutoFitColumns();
                }


                #endregion

                excelPackage.SaveAs(fileInfo);

                if (automatico)
                {
                    ss_pp.Update_Fecha_Fin("Facturación", "PendML", "Agrupado Totales");
                    EnvioCorreo_PdteWeb_PS_AT_TAM_v2(ruta_salida_archivo);
                }
                    

                

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ss_pp.Update_Comentario("Facturación", "PendML", "Agrupado Totales", e.Message);

            }
        }

        public void InformePendWeb_PSAT_TAM(string ruta_salida_archivo, bool automatico)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            
            int c = 1;
            int f = 1;


            DateTime fecha_actual = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime fecha_registro = new DateTime();
            int meses_pdtes = 0;
            int aniomes = 0;

            EndesaBusiness.facturacion.cuadro_mando.Sofisticados sof = 
                new facturacion.cuadro_mando.Sofisticados();
            
            Dictionary<string, EndesaEntity.facturacion.AgoraManual> agoraManual;

            bool tiene_complemento_a01 = false;
            bool sacar_resumen = false;
            bool sacar_detalle = true;
            bool sacar_portugal = true;

            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_totales;
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_agora;
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_Noagora;

            utilidades.Fechas utilfecha = new Fechas();

            try
            {
                               

                sof.Contruye_Sofisticados();
                agoraManual = CargaAgoraManual(DateTime.Now, DateTime.Now);

                FileInfo fileInfo = new FileInfo(ruta_salida_archivo);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);





                #region Resumen

                //var workSheet = excelPackage.Workbook.Worksheets.Add("Resumen");
                //var headerCells = workSheet.Cells[1, 1, 1, 17];
                //var headerFont = headerCells.Style.Font;

                //if (sacar_resumen)
                //{
                //    workSheet.Cells[1, 1].Value = "ÁGORA (SÍ/NO)";
                //    workSheet.Cells["A1:A1"].AutoFitColumns();
                //    workSheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    workSheet.Cells[1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(217, 225, 242));
                //    workSheet.Cells["A1:A2"].Merge = true;
                //    PintaRecuadro(excelPackage, "A1:A2");

                //    workSheet.Cells[3, 1].Value = "NO ÁGORA";
                //    workSheet.Cells[3, 1].Style.Font.Bold = true;
                //    workSheet.Cells["A3:A3"].AutoFitColumns();
                //    workSheet.Cells[3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    workSheet.Cells[3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells["A3:A22"].Merge = true;
                //    PintaRecuadro(excelPackage, "A3:A22");


                //    workSheet.Cells[1, 2].Value = "RESPONSABLE";
                //    workSheet.Cells[1, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    workSheet.Cells[1, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[1, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(217, 225, 242));
                //    workSheet.Cells["B1:B2"].Merge = true;
                //    PintaRecuadro(excelPackage, "B1:B2");

                //    workSheet.Cells[3, 2].Value = "Medida";
                //    workSheet.Cells[3, 2].Style.Font.Bold = true;
                //    workSheet.Cells[3, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    workSheet.Cells[3, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells["B3:B15"].Merge = true;
                //    PintaRecuadro(excelPackage, "B3:B15");

                //    workSheet.Cells[17, 2].Value = "Facturación";
                //    workSheet.Cells["B17:B17"].AutoFitColumns();
                //    workSheet.Cells[17, 2].Style.Font.Bold = true;
                //    workSheet.Cells[17, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    workSheet.Cells[17, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells["B17:B21"].Merge = true;
                //    PintaRecuadro(excelPackage, "B17:B21");

                //    workSheet.Cells[1, 3].Value = "SUBESTADO";
                //    workSheet.Cells[1, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    workSheet.Cells[1, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[1, 3].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(217, 225, 242));
                //    workSheet.Cells["C1:C2"].Merge = true;
                //    PintaRecuadro(excelPackage, "C1:C2");

                //    workSheet.Cells[2, 9].Value = "Evolución";
                //    workSheet.Cells[2, 9].Style.Font.Italic = true;
                //    workSheet.Cells[2, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    workSheet.Cells[2, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[2, 9].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[2, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 228, 214));
                //    PintaRecuadro(excelPackage, "I2:I2");

                //    workSheet.Cells[2, 12].Value = "Evolución";
                //    workSheet.Cells[2, 12].Style.Font.Italic = true;
                //    workSheet.Cells[2, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    workSheet.Cells[2, 12].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[2, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[2, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 228, 214));
                //    PintaRecuadro(excelPackage, "L2:L2");

                //    f = 3;
                //    c = 3;
                //    workSheet.Cells[f, c].Value = "01.A. Endesa - Telemedida";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "01.B. Endesa - TPL";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "01.D. NoEndesa - Telemedida";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "01.E. NoEndesa - TPL";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "02.A. El Punto no existe en el CS";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "03.A. CC Completa en el CS";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "04.A. CC Enviada a SCE ML";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "05.A. CC Rechazada por SCE ML";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "06.A. CC Incompleta  SCE ML";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "07.A. No Validada";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "07.B. Validada con Anomalías";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "07.C. Validada sin Anomalías";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "07.E.Modificada";
                //    PintaRecuadro(excelPackage, f, c); f++;

                //    workSheet.Cells[16, 2].Value = "Total Medidas";
                //    workSheet.Cells[16, 2].Style.Font.Bold = true;
                //    workSheet.Cells[16, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //    workSheet.Cells[16, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[16, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[16, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 242, 242));
                //    workSheet.Cells["B16:C16"].Merge = true;
                //    PintaRecuadro(excelPackage, "B16:C16");


                //    f = 17;
                //    c = 3;

                //    workSheet.Cells[f, c].Value = "08.A. Facturado en el día anterior";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "08.B. Bloqueado";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "08.C. Otros";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "09.A. El Punto está Extraído";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "10.A. Prefactura pendiente";
                //    PintaRecuadro(excelPackage, f, c); f++;


                //    workSheet.Cells[22, 2].Value = "Total Facturación";
                //    workSheet.Cells[22, 2].Style.Font.Bold = true;
                //    workSheet.Cells[22, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //    workSheet.Cells[22, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[22, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[22, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 242, 242));
                //    workSheet.Cells["B22:C22"].Merge = true;
                //    PintaRecuadro(excelPackage, "B22:C22");


                //    workSheet.Cells[1, 4].Value = "Pendiente PS";
                //    workSheet.Cells[1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    workSheet.Cells[1, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[1, 4].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(217, 225, 242));
                //    workSheet.Cells["D1:I1"].Merge = true;
                //    PintaRecuadro(excelPackage, "D1:I1");

                //    workSheet.Cells[1, 10].Value = "Pendiente Económico (M€)";
                //    workSheet.Cells[1, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    workSheet.Cells[1, 10].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[1, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[1, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(217, 225, 242));
                //    workSheet.Cells["J1:L1"].Merge = true;
                //    PintaRecuadro(excelPackage, "J1:L1");


                //    workSheet.Cells[23, 1].Value = "Total No Ágora";
                //    workSheet.Cells[23, 1].Style.Font.Bold = true;
                //    workSheet.Cells[23, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //    workSheet.Cells[23, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[23, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[23, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(226, 239, 218));
                //    workSheet.Cells["A23:C23"].Merge = true;
                //    PintaRecuadro(excelPackage, "A23:C23");


                //    workSheet.Cells[24, 2].Value = "Medida";
                //    workSheet.Cells[24, 2].Style.Font.Bold = true;
                //    workSheet.Cells[24, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    workSheet.Cells[24, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells["B24:B36"].Merge = true;
                //    PintaRecuadro(excelPackage, "B24:B36");

                //    workSheet.Cells[38, 2].Value = "Facturación";
                //    workSheet.Cells["B38:B38"].AutoFitColumns();
                //    workSheet.Cells[38, 2].Style.Font.Bold = true;
                //    workSheet.Cells[38, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    workSheet.Cells[38, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells["B38:B42"].Merge = true;
                //    PintaRecuadro(excelPackage, "B17:B21");



                //    workSheet.Cells[24, 1].Value = "SÍ ÁGORA";
                //    workSheet.Cells["A24:A24"].AutoFitColumns();
                //    workSheet.Cells[24, 1].Style.Font.Bold = true;
                //    workSheet.Cells[24, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    workSheet.Cells[24, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells["A24:A43"].Merge = true;
                //    PintaRecuadro(excelPackage, "A24:A43");



                //    f = 24;
                //    c = 3;
                //    workSheet.Cells[f, c].Value = "01.A. Endesa - Telemedida";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "01.B. Endesa - TPL";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "01.D. NoEndesa - Telemedida";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "01.E. NoEndesa - TPL";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "02.A. El Punto no existe en el CS";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "03.A. CC Completa en el CS";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "04.A. CC Enviada a SCE ML";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "05.A. CC Rechazada por SCE ML";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "06.A. CC Incompleta  SCE ML";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "07.A. No Validada";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "07.B. Validada con Anomalías";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "07.C. Validada sin Anomalías";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "07.E.Modificada";
                //    PintaRecuadro(excelPackage, f, c); f++;

                //    workSheet.Cells[37, 2].Value = "Total Medidas";
                //    workSheet.Cells[37, 2].Style.Font.Bold = true;
                //    workSheet.Cells[37, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //    workSheet.Cells[37, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[37, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[37, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 242, 242));
                //    workSheet.Cells["B37:C37"].Merge = true;
                //    PintaRecuadro(excelPackage, "B37:C37");


                //    f = 38;
                //    c = 3;

                //    workSheet.Cells[f, c].Value = "08.A. Facturado en el día anterior";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "08.B. Bloqueado";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "08.C. Otros";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "09.A. El Punto está Extraído";
                //    PintaRecuadro(excelPackage, f, c); f++;
                //    workSheet.Cells[f, c].Value = "10.A. Prefactura pendiente";
                //    PintaRecuadro(excelPackage, f, c); f++;


                //    workSheet.Cells[43, 2].Value = "Total Facturación";
                //    workSheet.Cells[43, 2].Style.Font.Bold = true;
                //    workSheet.Cells[43, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //    workSheet.Cells[43, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[43, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[43, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 242, 242));
                //    workSheet.Cells["B43:C43"].Merge = true;
                //    PintaRecuadro(excelPackage, "B43:C43");


                //    workSheet.Cells[44, 1].Value = "Total Ágora";
                //    workSheet.Cells[44, 1].Style.Font.Bold = true;
                //    workSheet.Cells[44, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //    workSheet.Cells[44, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[44, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[44, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(226, 239, 218));
                //    workSheet.Cells["A44:C44"].Merge = true;
                //    PintaRecuadro(excelPackage, "A44:C44");

                //    workSheet.Cells[45, 1].Value = "Total General";
                //    workSheet.Cells[45, 1].Style.Font.Bold = true;
                //    workSheet.Cells[45, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //    workSheet.Cells[45, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    workSheet.Cells[45, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[45, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(217, 225, 242));
                //    workSheet.Cells["A45:C45"].Merge = true;
                //    PintaRecuadro(excelPackage, "A45:C45");

                //    dic_totales = CargaTotales();
                //    dic_agora = CargaAgora();
                //    dic_Noagora = CargaNoAgora();



                //    bool firstOnly = true;
                //    int totales_dia = 0;
                //    int total_medida_no_agora = 0;
                //    int total_facturacion_no_agora = 0;
                //    foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> p in dic_Noagora)
                //    {
                //        totales_dia = 0;
                //        for (int i = 0; i < p.Value.Count; i++)
                //            totales_dia = totales_dia + p.Value[i].num_cups;


                //        if (firstOnly)
                //        {
                //            firstOnly = false;
                //            f = 3;
                //            c = 9;

                //            total_medida_no_agora = Total_Pendiente_DIFF(dic_agora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", null);

                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.B. Endesa - TPL");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                //            total_medida_no_agora += Total_Pendiente_DIFF(dic_Noagora, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "02. CC Rechazada por CS", null);

                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                //            total_medida_no_agora += Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "03. CC Completa en CS", null);

                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "03. CC Completa en CS", "03.A. CC Completa en el CS");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                //            total_medida_no_agora += Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "04. CC Enviada a SCE ML", null);

                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                //            total_medida_no_agora += Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "06. CC Incompleta  SCE ML", null);

                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                //            total_medida_no_agora += Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", null);

                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.A. No Validada");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.B. Validada con Anomalías");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.C. Validada sin Anomalías");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "07. LTP SCE", "07.E. Modificada");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                //            workSheet.Cells[f, c].Value = total_medida_no_agora;
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;



                //            total_facturacion_no_agora = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", null);

                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", "08.B. Bloqueado");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "08. El Punto no está Extraído", "08.C. Otros");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                //            total_facturacion_no_agora += Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "09. El Punto está Extraído", null);

                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                //            total_facturacion_no_agora += Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "10. Prefactura pendiente", null);

                //            workSheet.Cells[f, c].Value = Total_Pendiente_DIFF(dic_totales, p.Key, utilfecha.UltimoDiaHabilAnterior(p.Key), "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                //        }

                //    }

                //    headerCells = workSheet.Cells[1, 1, 1, 30];
                //    headerFont = headerCells.Style.Font;
                //    headerFont.Bold = true;
                //    var allCells = workSheet.Cells[1, 1, 50, 50];
                //    allCells.AutoFitColumns();

                //}

                #endregion




                #region Detalle

                var workSheet = excelPackage.Workbook.Worksheets.Add(DateTime.Now.ToString("Detalle"));
                    var headerCells = workSheet.Cells[1, 1, 1, 17];
                    var headerFont = headerCells.Style.Font;

                    workSheet.View.FreezePanes(2, 1);

                    headerFont.Bold = true;
                    workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                    workSheet.Cells[f, c].Value = "NIF"; c++;
                    workSheet.Cells[f, c].Value = "CLIENTE"; c++;
                    workSheet.Cells[f, c].Value = "FALTACONT"; c++;
                    workSheet.Cells[f, c].Value = "FPSERCON"; c++;
                    workSheet.Cells[f, c].Value = "CUPS13"; c++;
                    workSheet.Cells[f, c].Value = "CUPS20"; c++;
                    workSheet.Cells[f, c].Value = "TARIFA"; c++;
                    workSheet.Cells[f, c].Value = "CONTRATO"; c++;
                    workSheet.Cells[f, c].Value = "MES"; c++;
                    workSheet.Cells[f, c].Value = "DISTRIBUIDORA"; c++;
                    workSheet.Cells[f, c].Value = "ESTADO"; c++;
                    workSheet.Cells[f, c].Value = "SUBESTADO"; c++;
                    workSheet.Cells[f, c].Value = "MULTIMEDIDA"; c++;
                    workSheet.Cells[f, c].Value = "TAM"; c++;
                    workSheet.Cells[f, c].Value = "MESES PDTES FACTURAR"; c++;
                    workSheet.Cells[f, c].Value = "IMPORTE PDTE FACTURAR"; c++;
                    workSheet.Cells[f, c].Value = "ÁGORA";


                    for (int i = 1; i <= c; i++)
                    {
                        workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }


                    strSql = "SELECT empresa_titular, ps.NIF, ps.Cliente, ps.fAltaCont, ps.FPSERCON,"
                        + " p.cups13, ps.TARIFA, substr(ps.CUPS22, 1, 20) as CUPS20,"
                        + " contrato, mes, distribuidora, estado, subestado, multimedida, if (t.TAM IS NULL, 0, t.tam) AS TAM"
                        + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar_agrupado p"
                        + " LEFT OUTER JOIN cont.PS_AT ps ON"
                        + " ps.IDU = p.cups13"
                        + " LEFT OUTER JOIN fact.tam t ON"
                        + " t.CEMPTITU = p.empresa_titular AND"
                        + " t.CNIFDNIC = ps.NIF AND"
                        + " t.CUPS13 = ps.IDU"
                        + " WHERE p.empresa_titular <> 80"
                        // + " and p.cups13 <> 'VZZ0605131910'"
                        + " GROUP BY p.cups13";


                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        f++;
                        c = 1;
                        if (r["empresa_titular"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["empresa_titular"]);
                        c++;
                        if (r["NIF"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["NIF"].ToString();
                        c++;
                        if (r["Cliente"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["Cliente"].ToString();
                        c++;

                        if (r["fAltaCont"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["fAltaCont"]);
                        }

                        c++;
                        if (r["FPSERCON"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["FPSERCON"]);

                        c++;
                        if (r["cups13"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cups13"].ToString();

                        c++;
                        if (r["CUPS20"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CUPS20"].ToString();

                        c++;
                        if (r["TARIFA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                        c++;
                        if (r["contrato"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt64(r["contrato"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "###0";
                        }

                        c++;
                        if (r["mes"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["mes"]);
                            aniomes = Convert.ToInt32(r["mes"]);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                        }

                        c++;
                        if (r["distribuidora"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["distribuidora"].ToString();
                        c++;
                        if (r["estado"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["estado"].ToString();
                        c++;
                        if (r["subestado"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["subestado"].ToString();
                        c++;
                        if (r["multimedida"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["multimedida"].ToString();
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        c++;

                        if (r["TAM"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        }

                        c++;

                        if (r["mes"] != System.DBNull.Value)
                        {
                            //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                            workSheet.Cells[f, c].Value = meses_pdtes; c++;
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]) * meses_pdtes;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        }
                        else
                        {
                            c++; c++;
                        }

                        // AGORA
                        if (r["cups13"] != System.DBNull.Value)
                        {
                            tiene_complemento_a01 = sof.EsSofisticado(r["cups13"].ToString());
                            if (!tiene_complemento_a01)
                            {
                                if (r["CUPS20"] != System.DBNull.Value)
                                {
                                    EndesaEntity.facturacion.AgoraManual o;
                                    if (agoraManual.TryGetValue(r["CUPS20"].ToString(), out o))
                                        tiene_complemento_a01 = true;
                                }

                            }

                            if (tiene_complemento_a01 && Convert.ToInt32(r["empresa_titular"]) != 70)
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            else
                            {
                                workSheet.Cells[f, c].Value = "No";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                        }





                    }
                    db.CloseConnection();

                    headerCells = workSheet.Cells[1, 1, 1, c];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;
                    var allCells = workSheet.Cells[1, 1, f, c];

                    allCells.AutoFitColumns();

                    workSheet.View.FreezePanes(2, 1);
                    workSheet.Cells["A1:R1"].AutoFilter = true;
                    allCells.AutoFitColumns();
                
                #endregion

                #region 80_MT_Portugal
                if (sacar_portugal)
                {
                    c = 1;
                    f = 1;

                    workSheet = excelPackage.Workbook.Worksheets.Add("Portugal_MT");
                    headerCells = workSheet.Cells[1, 1, 1, 17];
                    headerFont = headerCells.Style.Font;


                    workSheet.View.FreezePanes(2, 1);

                    headerFont.Bold = true;
                    workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                    workSheet.Cells[f, c].Value = "NIF"; c++;
                    workSheet.Cells[f, c].Value = "CLIENTE"; c++;
                    workSheet.Cells[f, c].Value = "FALTACONT"; c++;
                    workSheet.Cells[f, c].Value = "FPSERCON"; c++;
                    workSheet.Cells[f, c].Value = "CUPS13"; c++;
                    workSheet.Cells[f, c].Value = "CUPS20"; c++;
                    workSheet.Cells[f, c].Value = "TARIFA"; c++;
                    workSheet.Cells[f, c].Value = "CONTRATO"; c++;
                    workSheet.Cells[f, c].Value = "MES"; c++;
                    workSheet.Cells[f, c].Value = "DISTRIBUIDORA"; c++;
                    workSheet.Cells[f, c].Value = "ESTADO"; c++;
                    workSheet.Cells[f, c].Value = "SUBESTADO"; c++;
                    workSheet.Cells[f, c].Value = "MULTIMEDIDA"; c++;
                    workSheet.Cells[f, c].Value = "TAM"; c++;
                    workSheet.Cells[f, c].Value = "MESES PDTES FACTURAR"; c++;
                    workSheet.Cells[f, c].Value = "IMPORTE PDTE FACTURAR"; c++;
                    workSheet.Cells[f, c].Value = "ÁGORA";

                    for (int i = 1; i <= c; i++)
                    {
                        workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }


                    DateTime ultima_carga = new DateTime();
                    strSql = "SELECT MAX(f_ult_mod) f_ult_mod FROM contratos_ps_mt";
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        ultima_carga = Convert.ToDateTime(r["f_ult_mod"]);
                    }
                    db.CloseConnection();

                    strSql = "SELECT empresa_titular, ps.faltacon, ps.CNIFDNIC, ps.DAPERSOC,"
                        + " p.cups13, ps.CTARIFA as tarifa, ps.CUPSREE as cups20,"
                        + " contrato, mes, distribuidora, estado, subestado, multimedida, if (t.TAM IS NULL, 0, t.tam) AS TAM"
                        + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar_agrupado p"
                        + " LEFT OUTER JOIN cont.contratos_ps_mt ps ON"
                        + " ps.CUPS = p.cups13"
                        + " LEFT OUTER JOIN fact.tam t ON"
                        //+ " t.CEMPTITU = p.empresa_titular AND"
                        //+ " t.CNIFDNIC = ps.NIF AND"
                        + " t.CUPS13 = ps.CUPS"
                        + " WHERE p.empresa_titular = 80 and"
                        + " ps.f_ult_mod >= '" + ultima_carga.ToString("yyyy-MM-dd") + "'"
                        + " GROUP BY p.cups13";


                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        f++;
                        c = 1;
                        if (r["empresa_titular"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["empresa_titular"]);
                        c++;

                        if (r["CNIFDNIC"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CNIFDNIC"].ToString();
                        c++;

                        if (r["DAPERSOC"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["DAPERSOC"].ToString();
                        c++;

                        if (r["faltacon"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["faltacon"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }

                        c++;

                        //if (r["FPSERCON"] != System.DBNull.Value)
                        //    workSheet.Cells[f, c].Value = Convert.ToInt32(r["FPSERCON"]);

                        c++;
                        if (r["cups13"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cups13"].ToString();

                        c++;
                        if (r["cups20"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cups20"].ToString();

                        c++;
                        if (r["TARIFA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                        c++;
                        if (r["contrato"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt64(r["contrato"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "###0";
                        }

                        c++;
                        if (r["mes"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["mes"]);
                            aniomes = Convert.ToInt32(r["mes"]);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                        }

                        c++;
                        if (r["distribuidora"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["distribuidora"].ToString();
                        c++;
                        if (r["estado"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["estado"].ToString();
                        c++;
                        if (r["subestado"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["subestado"].ToString();
                        c++;
                        if (r["multimedida"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["multimedida"].ToString();
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        c++;
                        if (r["TAM"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        }


                        c++;

                        if (r["mes"] != System.DBNull.Value)
                        {
                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                            workSheet.Cells[f, c].Value = meses_pdtes; c++;
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]) * meses_pdtes;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        }
                        else
                        {
                            c++; c++;
                        }


                        // AGORA
                        if (r["cups13"] != System.DBNull.Value)
                        {
                            tiene_complemento_a01 = sof.EsSofisticado(r["cups13"].ToString());
                            if (!tiene_complemento_a01)
                            {
                                if (r["cups20"] != System.DBNull.Value)
                                {
                                    EndesaEntity.facturacion.AgoraManual o;
                                    if (agoraManual.TryGetValue(r["CUPS20"].ToString(), out o))
                                        tiene_complemento_a01 = true;
                                }

                            }

                            if (tiene_complemento_a01)
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            else
                            {
                                workSheet.Cells[f, c].Value = "No";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }

                        }


                    }
                    db.CloseConnection();

                    headerCells = workSheet.Cells[1, 1, 1, c];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;
                    allCells = workSheet.Cells[1, 1, f, c];

                    allCells.AutoFitColumns();

                    workSheet.View.FreezePanes(2, 1);
                    workSheet.Cells["A1:R1"].AutoFilter = true;
                    allCells.AutoFitColumns();
                }


                #endregion#

                excelPackage.Save();

                if(automatico)
                    EnvioCorreo_PdteWeb_PS_AT_TAM(ruta_salida_archivo);
               
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);

            }
        }

        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaTotales()
        {
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>>();

  
            

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string estado = "";
            string subestado = "";
            int total_cups = 0;
            double total_tam = 0;
            DateTime fecha_informe = new DateTime();

            try
            {
                strSql = "SELECT t.fecha_informe, t.estado, t.subestado, t.num_cups, tam"
                    + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales_t t"
                    + " ORDER BY t.fecha_informe DESC, t.estado";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha_informe = Convert.ToDateTime(r["fecha_informe"]);
                    estado = r["estado"].ToString();
                    subestado = r["subestado"].ToString();
                    total_cups = Convert.ToInt32(r["num_cups"]);

                    if (r["tam"] != System.DBNull.Value)
                        total_tam = Convert.ToDouble(r["tam"]);
                    
                    List<EndesaEntity.medida.Pendiente_Totales> o;
                    if (!d.TryGetValue(fecha_informe, out o))
                    {
                        o = InicializaPendienteTotales();
                        d.Add(fecha_informe, o);
                    }                           


                    foreach(EndesaEntity.medida.Pendiente_Totales p in o)
                    {
                        if (p.estado == estado && p.subestado == subestado)
                        {
                            p.num_cups += total_cups;
                            p.tam += total_tam;
                        }
                            
                    }
                    

                }
                db.CloseConnection();
                return d;
            }
            catch(Exception e)
            {
                
                return null;
              
            }
        }
        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaTotales(List<int> lista_empresas)
        {
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>>();




            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string estado = "";
            string subestado = "";
            int total_cups = 0;
            double total_tam = 0;
            DateTime fecha_informe = new DateTime();

            try
            {
                strSql = "SELECT t.fecha_informe, t.estado, t.subestado, t.num_cups, tam"
                    + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales_t t"
                    + " where empresa_titular in (" + lista_empresas[0];

                for (int x = 1; x < lista_empresas.Count; x++)
                    strSql += "," + lista_empresas[x];

                strSql += ") ORDER BY t.fecha_informe DESC, t.estado";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha_informe = Convert.ToDateTime(r["fecha_informe"]);
                    estado = r["estado"].ToString();
                    subestado = r["subestado"].ToString();
                    total_cups = Convert.ToInt32(r["num_cups"]);

                    if (r["tam"] != System.DBNull.Value)
                        total_tam = Convert.ToDouble(r["tam"]);

                    List<EndesaEntity.medida.Pendiente_Totales> o;
                    if (!d.TryGetValue(fecha_informe, out o))
                    {
                        o = InicializaPendienteTotales();
                        d.Add(fecha_informe, o);
                    }


                    foreach (EndesaEntity.medida.Pendiente_Totales p in o)
                    {
                        if (p.estado == estado && p.subestado == subestado)
                        {
                            p.num_cups += total_cups;
                            p.tam += total_tam;
                        }

                    }


                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {

                return null;

            }
        }
        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaAgora(DateTime fd)
        {
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>>();




            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string estado = "";
            string subestado = "";
            int total_cups = 0;
            double tam = 0;
            DateTime fecha_informe = new DateTime();
            utilidades.Fechas utilFechas = new Fechas();            

            try
            {              

                strSql = "SELECT t.fecha_informe, t.estado, t.subestado, t.num_cups"
                    + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales_agora t"
                    + " where t.fecha_informe >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " ORDER BY t.fecha_informe, t.estado";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                                       

                    fecha_informe = Convert.ToDateTime(r["fecha_informe"]);
                    estado = r["estado"].ToString();
                    subestado = r["subestado"].ToString();
                    total_cups = Convert.ToInt32(r["num_cups"]);

                    List<EndesaEntity.medida.Pendiente_Totales> o;
                    if (!d.TryGetValue(fecha_informe, out o))
                    {
                        o = InicializaPendienteTotales();
                        d.Add(fecha_informe, o);
                    }


                    foreach (EndesaEntity.medida.Pendiente_Totales p in o)
                    {
                        if (p.estado == estado && p.subestado == subestado)
                        {
                            p.num_cups += total_cups;                            
                        }
                            
                    }


                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {

                return null;

            }
        }
        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaAgora(DateTime fd, List<int> lista_empresas)
        {
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>>();




            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string estado = "";
            string subestado = "";
            int total_cups = 0;
            double tam = 0;
            DateTime fecha_informe = new DateTime();
            utilidades.Fechas utilFechas = new Fechas();

            try
            {

                strSql = "SELECT t.fecha_informe, t.estado, t.subestado, t.num_cups"
                    + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales_agora t"
                    + " where t.fecha_informe >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " and empresa_titular in (" + lista_empresas[0];

                for (int x = 1; x < lista_empresas.Count; x++)
                    strSql = "," + lista_empresas[x];

                strSql +=  ") ORDER BY t.fecha_informe, t.estado";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {


                    fecha_informe = Convert.ToDateTime(r["fecha_informe"]);
                    estado = r["estado"].ToString();
                    subestado = r["subestado"].ToString();
                    total_cups = Convert.ToInt32(r["num_cups"]);

                    List<EndesaEntity.medida.Pendiente_Totales> o;
                    if (!d.TryGetValue(fecha_informe, out o))
                    {
                        o = InicializaPendienteTotales();
                        d.Add(fecha_informe, o);
                    }


                    foreach (EndesaEntity.medida.Pendiente_Totales p in o)
                    {
                        if (p.estado == estado && p.subestado == subestado)
                        {
                            p.num_cups += total_cups;
                        }

                    }


                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {

                return null;

            }
        }
        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaAgora_TAM(DateTime fd, List<int> lista_empresas)
        {
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>>();




            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string estado = "";
            string subestado = "";
            int total_cups = 0;
            double tam = 0;
            DateTime fecha_informe = new DateTime();
            utilidades.Fechas utilFechas = new Fechas();

            try
            {

                strSql = "SELECT t.fecha_informe, t.estado, t.subestado, t.num_cups, tam"
                    + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales_agora_t t"
                    + " where empresa_titular in (" + lista_empresas[0];

                for (int x = 1; x < lista_empresas.Count; x++)
                    strSql += "," + lista_empresas[x];

                strSql += ") and t.fecha_informe >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " ORDER BY t.fecha_informe, t.estado";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {


                    fecha_informe = Convert.ToDateTime(r["fecha_informe"]);
                    estado = r["estado"].ToString();
                    subestado = r["subestado"].ToString();
                    total_cups = Convert.ToInt32(r["num_cups"]);

                    if (r["tam"] != System.DBNull.Value)
                        tam = Convert.ToDouble(r["tam"]);

                    List<EndesaEntity.medida.Pendiente_Totales> o;
                    if (!d.TryGetValue(fecha_informe, out o))
                    {
                        o = InicializaPendienteTotales();
                        d.Add(fecha_informe, o);
                    }


                    foreach (EndesaEntity.medida.Pendiente_Totales p in o)
                    {
                        if (p.estado == estado && p.subestado == subestado)
                        {
                            p.num_cups += total_cups;
                            p.tam += tam;
                        }

                    }


                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {

                return null;

            }
        }
        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaNoAgora(DateTime fd)
        {
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>>();

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string estado = "";
            string subestado = "";
            int total_cups = 0;
            double tam = 0;
            DateTime fecha_informe = new DateTime();
            utilidades.Fechas utilFechas = new Fechas();

            try
            {


                strSql = "SELECT t.fecha_informe, t.estado, t.subestado, t.num_cups"
                    + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales_noagora t"
                    + " where t.fecha_informe >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " ORDER BY t.fecha_informe, t.estado";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha_informe = Convert.ToDateTime(r["fecha_informe"]);
                    estado = r["estado"].ToString();
                    subestado = r["subestado"].ToString();
                    total_cups = Convert.ToInt32(r["num_cups"]);

                   

                    List<EndesaEntity.medida.Pendiente_Totales> o;
                    if (!d.TryGetValue(fecha_informe, out o))
                    {
                        o = InicializaPendienteTotales();
                        d.Add(fecha_informe, o);
                    }


                    foreach (EndesaEntity.medida.Pendiente_Totales p in o)
                    {
                        if (p.estado == estado && p.subestado == subestado)
                        {
                            p.num_cups += total_cups;
                            
                        }

                    }


                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {

                return null;

            }
        }
        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaNoAgora(DateTime fd, List<int> lista_empresas)
        {
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>>();

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string estado = "";
            string subestado = "";
            int total_cups = 0;
            double tam = 0;
            DateTime fecha_informe = new DateTime();
            utilidades.Fechas utilFechas = new Fechas();

            try
            {


                strSql = "SELECT t.fecha_informe, t.estado, t.subestado, t.num_cups"
                    + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales_noagora t"
                     + " where empresa_titular in (" + lista_empresas[0];

                for (int x = 1; x < lista_empresas.Count; x++)
                    strSql += "," + lista_empresas[x];

                strSql += ") and t.fecha_informe >= '" + fd.ToString("yyyy-MM-dd") + "'"                    
                    + " ORDER BY t.fecha_informe, t.estado";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha_informe = Convert.ToDateTime(r["fecha_informe"]);
                    estado = r["estado"].ToString();
                    subestado = r["subestado"].ToString();
                    total_cups = Convert.ToInt32(r["num_cups"]);



                    List<EndesaEntity.medida.Pendiente_Totales> o;
                    if (!d.TryGetValue(fecha_informe, out o))
                    {
                        o = InicializaPendienteTotales();
                        d.Add(fecha_informe, o);
                    }


                    foreach (EndesaEntity.medida.Pendiente_Totales p in o)
                    {
                        if (p.estado == estado && p.subestado == subestado)
                        {
                            p.num_cups = total_cups;

                        }

                    }


                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {

                return null;

            }
        }
        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaNoAgora_TAM(DateTime fd, List<int> lista_empresas)
        {
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>>();

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string estado = "";
            string subestado = "";
            int total_cups = 0;
            double tam = 0;
            DateTime fecha_informe = new DateTime();
            utilidades.Fechas utilFechas = new Fechas();

            try
            {


                strSql = "SELECT t.fecha_informe, t.estado, t.subestado, t.num_cups, tam"
                    + " FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_totales_noagora_t t"
                    + " where empresa_titular in (" + lista_empresas[0];

                for (int x = 1; x < lista_empresas.Count; x++)
                    strSql += "," + lista_empresas[x];

                strSql += ") and t.fecha_informe >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " ORDER BY t.fecha_informe, t.estado";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha_informe = Convert.ToDateTime(r["fecha_informe"]);
                    estado = r["estado"].ToString();
                    subestado = r["subestado"].ToString();
                    total_cups = Convert.ToInt32(r["num_cups"]);

                    if (r["tam"] != System.DBNull.Value)
                        tam = Convert.ToDouble(r["tam"]);

                    List<EndesaEntity.medida.Pendiente_Totales> o;
                    if (!d.TryGetValue(fecha_informe, out o))
                    {
                        o = InicializaPendienteTotales();
                        d.Add(fecha_informe, o);
                    }


                    foreach (EndesaEntity.medida.Pendiente_Totales p in o)
                    {
                        if (p.estado == estado && p.subestado == subestado)
                        {
                            p.num_cups = total_cups;
                            p.tam = tam;
                        }
                            
                    }


                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {

                return null;

            }
        }

        private List<EndesaEntity.medida.Pendiente_Totales> InicializaPendienteTotales()
        {
            List<EndesaEntity.medida.Pendiente_Totales> t = new List<EndesaEntity.medida.Pendiente_Totales>();
            EndesaEntity.medida.Pendiente_Totales c = new EndesaEntity.medida.Pendiente_Totales();

            c.estado = "01. Pendiente de medida";
            c.subestado = "01.A. Endesa - Telemedida";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "01. Pendiente de medida";
            c.subestado = "01.B. Endesa - TPL";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "01. Pendiente de medida";
            c.subestado = "01.D. NoEndesa - Telemedida";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "01. Pendiente de medida";
            c.subestado = "01.E. NoEndesa - TPL";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "02. CC Rechazada por CS";
            c.subestado = "02.A. El Punto no existe en el CS";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "03. CC Completa en CS";
            c.subestado = "03.A. CC Completa en el CS";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "04. CC Enviada a SCE ML";
            c.subestado = "04.A. CC Enviada a SCE ML";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "06. CC Incompleta  SCE ML";
            c.subestado = "06.A. CC Incompleta  SCE ML";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "07. LTP SCE";
            c.subestado = "07.A. No Validada";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "07. LTP SCE";
            c.subestado = "07.B. Validada con Anomalías";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "07. LTP SCE";
            c.subestado = "07.C. Validada sin Anomalías";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "07. LTP SCE";
            c.subestado = "07.E. Modificada";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "08. El Punto no está Extraído";
            c.subestado = "08.A. Facturado en el día anterior";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "08. El Punto no está Extraído";
            c.subestado = "08.B. Bloqueado";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "08. El Punto no está Extraído";
            c.subestado = "08.C. Otros";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "09. El Punto está Extraído";
            c.subestado = "09.A. El Punto está Extraído";
            t.Add(c);

            c = new EndesaEntity.medida.Pendiente_Totales();
            c.estado = "10. Prefactura pendiente";
            c.subestado = "10.A. Prefactura pendiente";
            t.Add(c);

            return t;
        }

        private int Total_Pendiente(Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic, DateTime fecha, 
            string estado, string subestado)
        {
            int total = 0;

            List<EndesaEntity.medida.Pendiente_Totales> o;
            if(dic.TryGetValue(fecha, out o))
            {
                for(int i = 0; i < o.Count; i++)
                {
                    if (o[i].estado == estado && (o[i].subestado == subestado || subestado == null))
                        total = total + o[i].num_cups;

                }
            }

            return total;        
        }

        private double Total_Pendiente_TAM(Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic, DateTime fecha,
            string estado, string subestado)
        {
            double total = 0;

            List<EndesaEntity.medida.Pendiente_Totales> o;
            if (dic.TryGetValue(fecha, out o))
            {
                for (int i = 0; i < o.Count; i++)
                {
                    if (o[i].estado == estado && (o[i].subestado == subestado || subestado == null))
                        total = total + o[i].tam;
                }
            }

            return total;
        }

        private int Total_Pendiente_DIFF(Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic, DateTime fecha1,
            DateTime fecha2, string estado, string subestado)
        {

            // fecha2 suele ser un dia menos
            // o ultimo_dia_habil -1

            int total1 = 0;
            int total2 = 0;
            

            List<EndesaEntity.medida.Pendiente_Totales> o;
            if (dic.TryGetValue(fecha1.Date, out o))
            {
                for (int i = 0; i < o.Count; i++)
                {
                    if (o[i].estado == estado && (o[i].subestado == subestado || subestado == null))
                        total1 = total1 + o[i].num_cups;

                }
            }

            if (dic.TryGetValue(fecha2.Date, out o))
            {
                for (int i = 0; i < o.Count; i++)
                {
                    if (o[i].estado == estado && (o[i].subestado == subestado || subestado == null))
                        total2 = total2 + o[i].num_cups;

                }
            }

            return total2 - total1;
        }

        private double Total_Pendiente_DIFF_TAM(Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic, DateTime fecha1,
            DateTime fecha2, string estado, string subestado)
        {

            // fecha2 suele ser un dia menos
            // o ultimo_dia_habil -1

            double total1 = 0;
            double total2 = 0;


            List<EndesaEntity.medida.Pendiente_Totales> o;
            if (dic.TryGetValue(fecha1.Date, out o))
            {
                for (int i = 0; i < o.Count; i++)
                {
                    if (o[i].estado == estado && (o[i].subestado == subestado || subestado == null))
                        total1 = total1 + o[i].tam;

                }
            }

            if (dic.TryGetValue(fecha2.Date, out o))
            {
                for (int i = 0; i < o.Count; i++)
                {
                    if (o[i].estado == estado && (o[i].subestado == subestado || subestado == null))
                        total2 = total2 + o[i].tam;

                }
            }

            return total2 - total1;
        }

        private int Total_Pdte_DIA(Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic, DateTime fecha)
        {
            int total_dia = 0;
            List<EndesaEntity.medida.Pendiente_Totales> o;
            if(dic.TryGetValue(fecha.Date, out o))
            {
                for (int i = 0; i < o.Count; i++)
                    total_dia = total_dia + o[i].num_cups;
            }

            return total_dia;
        }

        private double Total_Pdte_DIA_TAM(Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic, DateTime fecha)
        {
            double total_dia_tam = 0;
            List<EndesaEntity.medida.Pendiente_Totales> o;
            if (dic.TryGetValue(fecha.Date, out o))
            {
                for (int i = 0; i < o.Count; i++)
                    total_dia_tam = total_dia_tam + o[i].tam;
            }

            return total_dia_tam;
        }

        private void PintaRecuadro_Total(ExcelPackage excelPackage, string posicion)
        {
            var workSheet = excelPackage.Workbook.Worksheets.First();
            workSheet.Cells[posicion].Style.Border.Top.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[posicion].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[posicion].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[posicion].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
        }
               
        public void Crea_Pdte_Agora()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int i = 0;
            int k = 0;
            int totalRegistros = 0;

            try
            {
                complementosContratoPS = new contratacion.ComplementosContrato(DateTime.Now, DateTime.Now, "A01");
                agoraPortugal = new contratacion.Agora_Portugal();

                strSql = "DELETE FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_agora";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "DELETE FROM dt_vw_ed_f_detalle_pendiente_facturar_agrupado_agora_t";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "SELECT empresa_titular, cups13, contrato, mes, distribuidora, estado, subestado, multimedida, fecha_informe"                    
                     + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar_agrupado";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    i++;
                    k++;
                    if (firstOnly)
                    {
                        firstOnly = false;
                        sb.Append("replace into dt_vw_ed_f_detalle_pendiente_facturar_agrupado_agora");
                        sb.Append(" (empresa_titular,cups13,contrato,mes,distribuidora,estado,subestado,");
                        sb.Append("multimedida,agora,fecha_informe) values");
                    }

                    #region Campos
                    if (r["empresa_titular"] != System.DBNull.Value)
                        sb.Append(" (").Append(Convert.ToInt32(r["empresa_titular"])).Append(",");
                    else
                        sb.Append("(null,");

                    if (r["cups13"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cups13"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["contrato"] != System.DBNull.Value)
                        sb.Append(r["contrato"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["mes"] != System.DBNull.Value)
                        sb.Append(Convert.ToInt32(r["mes"])).Append(",");
                    else
                        sb.Append("null,");

                    if (r["distribuidora"] != System.DBNull.Value)
                        sb.Append("'").Append(FuncionesTexto.RT(r["distribuidora"].ToString())).Append("',");
                    else
                        sb.Append("null,");

                    if (r["estado"] != System.DBNull.Value)
                        sb.Append("'").Append(r["estado"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["subestado"] != System.DBNull.Value)
                        sb.Append("'").Append(r["subestado"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["multimedida"] != System.DBNull.Value)
                        sb.Append("'").Append(r["multimedida"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    // AGORA

                    if (complementosContratoPS.TieneComplemento(r["cups13"].ToString()))
                        sb.Append("'").Append("S").Append("',");
                    else if(agoraPortugal.EsAgora(r["cups13"].ToString()))
                        sb.Append("'").Append("S").Append("',");
                    else
                        sb.Append("'").Append("N").Append("',");

                    if (r["fecha_informe"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fecha_informe"]).ToString("yyyy-MM-dd")).Append("'),");
                    else
                        sb.Append("null),");

                    #endregion

                    if (i == 500)
                    {
                        Console.WriteLine("Anexamos " + k + " / " + totalRegistros);
                        firstOnly = true;
                        db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        i = 0;
                    }

                }

                if (i > 0)
                {
                    Console.WriteLine("Anexamos " + k + " / " + totalRegistros);
                    firstOnly = true;
                    db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    i = 0;
                }
                db.CloseConnection();

                



                firstOnly = true;
                strSql = "SELECT empresa_titular, cups13, contrato, mes, distribuidora, estado, subestado, multimedida, fecha_informe, tam"
                     + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar_agrupado_t";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    i++;
                    k++;
                    if (firstOnly)
                    {
                        firstOnly = false;
                        sb.Append("replace into dt_vw_ed_f_detalle_pendiente_facturar_agrupado_agora_t");
                        sb.Append(" (empresa_titular,cups13,contrato,mes,distribuidora,estado,subestado,");
                        sb.Append("multimedida,agora,fecha_informe, tam) values");
                    }

                    #region Campos
                    if (r["empresa_titular"] != System.DBNull.Value)
                        sb.Append(" (").Append(Convert.ToInt32(r["empresa_titular"])).Append(",");
                    else
                        sb.Append("(null,");

                    if (r["cups13"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cups13"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["contrato"] != System.DBNull.Value)
                        sb.Append(r["contrato"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["mes"] != System.DBNull.Value)
                        sb.Append(Convert.ToInt32(r["mes"])).Append(",");
                    else
                        sb.Append("null,");

                    if (r["distribuidora"] != System.DBNull.Value)
                        sb.Append("'").Append(FuncionesTexto.RT(r["distribuidora"].ToString())).Append("',");
                    else
                        sb.Append("null,");

                    if (r["estado"] != System.DBNull.Value)
                        sb.Append("'").Append(r["estado"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["subestado"] != System.DBNull.Value)
                        sb.Append("'").Append(r["subestado"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["multimedida"] != System.DBNull.Value)
                        sb.Append("'").Append(r["multimedida"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (complementosContratoPS.TieneComplemento(r["cups13"].ToString()))
                        sb.Append("'").Append("S").Append("',");
                    else if (agoraPortugal.EsAgora(r["cups13"].ToString()))
                        sb.Append("'").Append("S").Append("',");
                    else
                        sb.Append("'").Append("N").Append("',");

                    if (r["fecha_informe"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fecha_informe"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["tam"] != System.DBNull.Value)
                        sb.Append(r["tam"].ToString().Replace(",", ".")).Append("),");
                    else
                        sb.Append("null),");


                    #endregion

                    if (i == 500)
                    {
                        Console.WriteLine("Anexamos " + k + " / " + totalRegistros);
                        firstOnly = true;
                        db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        i = 0;
                    }

                }

                if (i > 0)
                {
                    Console.WriteLine("Anexamos " + k + " / " + totalRegistros);
                    firstOnly = true;
                    db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    i = 0;
                }
                db.CloseConnection();

            }
            catch(Exception e)
            {
                ficheroLog.addError("Crea_Pdte_Agora: " + e.Message);
            }
        }

        private Dictionary<string, EndesaEntity.facturacion.AgoraManual> CargaAgoraManual(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, EndesaEntity.facturacion.AgoraManual> d
                = new Dictionary<string, EndesaEntity.facturacion.AgoraManual>();

            try
            {
                strSql = "Select Empresa, NIF, CCOUNIPS, CUPS20, DAPERSOC"
                    + " from fact.cm_sofisticados "
                    + " where (FD <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FH >= '" + fh.ToString("yyyy-MM-dd") + "' or FH is null)";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.AgoraManual c = new EndesaEntity.facturacion.AgoraManual();
                    c.empresa = r["Empresa"].ToString();
                    c.nif = r["NIF"].ToString();
                    c.cups13 = r["CCOUNIPS"].ToString();
                    c.cups20 = r["CUPS20"].ToString();
                    c.cliente = r["DAPERSOC"].ToString();

                    EndesaEntity.facturacion.AgoraManual o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        private void PintaRecuadro(ExcelPackage excelPackage, string posicion)
        {
            var workSheet = excelPackage.Workbook.Worksheets.First();
            workSheet.Cells[posicion].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[posicion].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[posicion].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[posicion].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

        private void PintaRecuadro(ExcelPackage excelPackage, int f, int c)
        {
            var workSheet = excelPackage.Workbook.Worksheets.First();
            workSheet.Cells[f, c].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[f, c].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[f, c].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }



    }
}

