using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class SCEA
    {
        utilidades.Param param;
        Dictionary<string, EndesaEntity.medida.SCEA_Tabla> dic;
        List<EndesaEntity.medida.SCEA_Tabla> lista_informe;
        logs.Log ficheroLog;

        public SCEA()
        {
            lista_informe = new List<EndesaEntity.medida.SCEA_Tabla>();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "SCEA");
            param = new utilidades.Param("scea_param", servidores.MySQLDB.Esquemas.MED);
            dic = new Dictionary<string, EndesaEntity.medida.SCEA_Tabla>();
        }

        public void CargaMultipuntos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "SELECT s.IDU FROM scea s"
                    + " WHERE s.N_PMs_CS > 1"
                    + " GROUP BY s.IDU; ";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.SCEA_Tabla c = new EndesaEntity.medida.SCEA_Tabla();
                    c.cup13 = r["IDU"].ToString();                                       
                    dic.Add(c.cup13, c);

                }
                db.CloseConnection();


            }
            catch(Exception e)
            {

            }
        }

        public bool Es_Multipunto(string cups13)
        {
            EndesaEntity.medida.SCEA_Tabla o;
            return dic.TryGetValue(cups13, out o);
        }

        public void Informe_Estado_Pdte_Origen_SCEA()
        {

            string archivo_informe = "";
            bool actualizado_scea = false;

            EndesaBusiness.utilidades.FechasProcesos fp = new utilidades.FechasProcesos();

            try
            {

                actualizado_scea = ActualizadoSCEA();

                ficheroLog.Add("Actualizado Pendiente: " + (actualizado_scea ? "S" : "N"));
                Console.WriteLine("Actualizado Pendiente: " + (actualizado_scea ? "S" : "N"));
                ficheroLog.Add("Última actualización Informe: " + fp.GetFechaProceso("INF_ESTADO_PDTE_ORIGEN_SCEA").ToString("dd/MM/yyyy HH:mm:ss"));
                Console.WriteLine("Última actualización Informe: " + fp.GetFechaProceso("INF_ESTADO_PDTE_ORIGEN_SCEA").ToString("dd/MM/yyyy HH:mm:ss"));

                if (actualizado_scea && (fp.GetFechaProceso("INF_ESTADO_PDTE_ORIGEN_SCEA").Date < DateTime.Now.Date))
                {
                    EndesaBusiness.utilidades.Fichero.BorrarArchivos_MenosUltimo(param.GetValue("salida_programada"), param.GetValue("prefijo_archivo"), "xlsx", null);
                    archivo_informe = GeneraInforme_Estado_Pdte_Origen_SCEA(RecopilaDatos_Estado_Pdte_Origen_SCEA());
                    if (param.GetValue("mail_enviar") == "S")
                        EnvioCorreo(archivo_informe);

                    fp.Update("INF_ESTADO_PDTE_ORIGEN_SCEA", DateTime.Now);

                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }


        private bool ActualizadoSCEA()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime fecha = new DateTime();

            try
            {

                strSql = "SELECT MAX(F_ULT_MOD) max_fecha FROM scea";
                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha = Convert.ToDateTime(r["max_fecha"]);
                }
                db.CloseConnection();               

                if (fecha.Date == DateTime.Now.Date)
                    return true;
                else
                    return false;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("ActualizadoSCEA: " + e.Message);
                return false;
            }
        }

        private List<EndesaEntity.medida.SCEA_Tabla> RecopilaDatos_Estado_Pdte_Origen_SCEA()
        {
            

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            List<EndesaEntity.medida.SCEA_Tabla> lista =
                new List<EndesaEntity.medida.SCEA_Tabla>();



            try
            {               



                strSql = "SELECT IDU, CUPS20, ET," 
                    + " if(FA = '0000-00-00', null, FA) as FA," 
                    + " if(FprevB = '0000-00-00', null, FprevB) as FprevB,"
                    + "Tarifa, Descripcion_Estado_SCE,"
                    + " CLIENTE, NIF, c3, `End`, tdistri, ddistrib, Tipo_TOP, Tipo_PM_Calculado," 
                    + " PotMaxCont, PROVINC, Coment_SCE, fComentSCE, Info, NOMBRE_GESTOR, TERRITORIAL," 
                    + " RESPONSABLE, PMs_PDTES, aaaammTrabajo, diaHabil, PrimerMesPDTE, Estado," 
                    + " Subestado, nMesesPDTES, Clase, usuario, Grupo_Resolucion, aaaammLTP, Definicion," 
                    + " if(fhLTP = '0000-00-00', null, fhLTP) as fhLTP,"
                    + " Anomalia, dificultad, t, ActCCR, ReactCCR, tipoIncompletitud, DescripcionIncompletitud," 
                    + " Incompletitudes, Existe_FactD, "
                    + " if(fdesdeCCR = '0000-00-00', null, fdesdeCCR) as fdesdeCCR,"                    
                    + " if(fhastaCCR = '0000-00-00', null, fhastaCCR) as fhastaCCR,"                     
                    + " PotMaxCCR,"
                    + " if(Maxdefh = '0000-00-00', null, Maxdefh) as Maxdefh," 
                    + " Gestion_ATR_Propia,"
                    + " AFactD, ExecesosPotencia, ExcesosReactiva, esPrimeraFactura, aaaamm_ULT_FACT,"
                    + " if(fdULTFACT = '0000-00-00', null, fdULTFACT) as fdULTFACT,"
                    + " if(fhULTFACT = '0000-00-00', null, fhULTFACT) as fhULTFACT,"                    
                    + " A_ULT_FACT, MinA, MaxA, MedA, Peor_fuente, ULT_Fuente, TOTAL_DEUDA,"
                    + " if(ult_recep_web = '0000-00-00', null, ult_recep_web) as ult_recep_web,"                    
                    + " Grupo, Ranking_TAM, Ranking_MaxFact, TAM_por_CUPS, factMax, "
                    + " diaFactMed, diaFactMax, N_PMs_CS, TLFCS, TM_OK, CONSUMO_0, PERDIDAS_ML, "
                    + " if(Maxdefh_Absoluta = '0000-00-00', null, Maxdefh_Absoluta) as Maxdefh_Absoluta,"                    
                    + " AAAAMM_Maxdefh,"
                    + " AAAAMM_Maxdefh_Absoluta, "
                    + " if(Fecha_Baja = '0000-00-00', null, Fecha_Baja) as Fecha_Baja,"                    
                    + " Proceso_Concursal,"
                    + " MAXDEFH_CIERRE, AAAAMM_MAXDEFH_CIERRE, F_ULT_MOD"
                    + " FROM med.scea";

                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {                   

                    EndesaEntity.medida.SCEA_Tabla t = new EndesaEntity.medida.SCEA_Tabla();

                    //1
                    if (r["IDU"] != System.DBNull.Value)                    
                        t.cup13 = r["IDU"].ToString();                     

                    if (r["CUPS20"] != System.DBNull.Value)
                        t.cups20 = r["CUPS20"].ToString();                    

                    if (r["ET"] != System.DBNull.Value)                    
                        t.empresa = r["ET"].ToString();

                    if (r["FA"] != System.DBNull.Value)
                        t.fecha_anexion = Convert.ToDateTime(r["FA"]);

                    if (r["FprevB"] != System.DBNull.Value)
                        t.fecha_prevista_baja = Convert.ToDateTime(r["FprevB"]);

                    if (r["Tarifa"] != System.DBNull.Value)
                        t.tarifa = r["Tarifa"].ToString();

                    if (r["Descripcion_Estado_SCE"] != System.DBNull.Value)
                        t.estado_contrato = r["Descripcion_Estado_SCE"].ToString();

                    if (r["CLIENTE"] != System.DBNull.Value)
                       t.cliente = r["CLIENTE"].ToString();

                    if (r["NIF"] != System.DBNull.Value)
                        t.nif = r["NIF"].ToString();

                    //10
                    if (r["c3"] != System.DBNull.Value)
                        t.c3 = r["c3"].ToString();

                    if (r["End"] != System.DBNull.Value)
                        t.end = r["End"].ToString();

                    if (r["tdistri"] != System.DBNull.Value)
                        t.tipo_distribora = r["tdistri"].ToString();

                    if (r["ddistrib"] != System.DBNull.Value)
                        t.descripcion_distribuidora = r["ddistrib"].ToString();

                    if (r["Tipo_TOP"] != System.DBNull.Value)
                        t.tipo_top = r["Tipo_TOP"].ToString();

                    if (r["Tipo_PM_Calculado"] != System.DBNull.Value)
                        t.tipo_pm_calculado = Convert.ToInt32(r["Tipo_PM_Calculado"]);

                    if(r["PotMaxCont"] != System.DBNull.Value)
                        t.potencia_maxima_contratada = Convert.ToDouble(r["PotMaxCont"]);

                    if (r["PROVINC"] != System.DBNull.Value)
                        t.provincia = r["PROVINC"].ToString();

                    if(r["Coment_SCE"] != System.DBNull.Value)
                        t.comentario_sce = r["Coment_SCE"].ToString();
                    
                    if (r["fComentSCE"] != System.DBNull.Value)
                        t.fecha_comentario_sce = Convert.ToDateTime(r["fComentSCE"]);

                    //20
                    if (r["Info"] != System.DBNull.Value)
                        t.info = r["Info"].ToString();

                    if(r["NOMBRE_GESTOR"] != System.DBNull.Value)
                        t.nombre_gestor = r["NOMBRE_GESTOR"].ToString();

                    if (r["TERRITORIAL"] != System.DBNull.Value)
                       t.territorial = r["TERRITORIAL"].ToString();

                    if (r["RESPONSABLE"] != System.DBNull.Value)
                        t.responsable = r["RESPONSABLE"].ToString();

                    if (r["PMs_PDTES"] != System.DBNull.Value)
                       t.pm_pdtes = Convert.ToInt32(r["PMs_PDTES"]);

                    if(r["aaaammTrabajo"] != System.DBNull.Value)
                        t.aaaamm_trabajo = Convert.ToInt32(r["aaaammTrabajo"]);

                    if(r["diaHabil"] != System.DBNull.Value)
                        t.dia_habil = Convert.ToInt32(r["diaHabil"]);

                    if (r["PrimerMesPDTE"] != System.DBNull.Value)
                        t.primer_mes_pdte = Convert.ToInt32(r["PrimerMesPDTE"]);

                    if (r["Estado"] != System.DBNull.Value)
                        t.estado = r["Estado"].ToString();

                    if (r["Subestado"] != System.DBNull.Value)
                        t.subestado = r["Subestado"].ToString();

                    //30
                    if (r["nMesesPDTES"] != System.DBNull.Value)
                        t.num_meses_pdtes = Convert.ToInt32(r["nMesesPDTES"]);

                    if (r["Clase"] != System.DBNull.Value)
                       t.clase = Convert.ToInt32(r["Clase"]);

                    if (r["usuario"] != System.DBNull.Value)
                        t.usuario = r["usuario"].ToString();

                    if (r["Grupo_Resolucion"] != System.DBNull.Value)
                        t.grupo_resolucion = r["Grupo_Resolucion"].ToString();

                    if (r["aaaammLTP"] != System.DBNull.Value)
                        t.aaaammlpt = Convert.ToInt32(r["aaaammLTP"]);

                    if (r["Definicion"] != System.DBNull.Value)
                        t.definicion = r["Definicion"].ToString();

                    if (r["fhLTP"] != System.DBNull.Value)
                        t.fechahasta_ltp = Convert.ToDateTime(r["fhLTP"]);

                    if (r["Anomalia"] != System.DBNull.Value)
                        t.anomalia = r["Anomalia"].ToString();

                    if (r["dificultad"] != System.DBNull.Value)
                        t.dificultad = Convert.ToInt32(r["dificultad"]);

                    if (r["t"] != System.DBNull.Value)
                        t.t = r["t"].ToString();

                    //40
                    if (r["ActCCR"] != System.DBNull.Value)
                        t.activa_ccr = Convert.ToInt32(r["ActCCR"]);

                    if (r["ReactCCR"] != System.DBNull.Value)
                        t.reactiva_ccr = Convert.ToInt32(r["ReactCCR"]);

                    if (r["tipoIncompletitud"] != System.DBNull.Value)
                        t.tipo_incompletitud = r["tipoIncompletitud"].ToString();

                    if (r["DescripcionIncompletitud"] != System.DBNull.Value)
                        t.descripcion_incompletitud = r["DescripcionIncompletitud"].ToString();

                    if (r["Incompletitudes"] != System.DBNull.Value)
                       t.incompletitudes = Convert.ToInt32(r["Incompletitudes"]);

                    if (r["Existe_FactD"] != System.DBNull.Value)
                        t.existe_FactD = r["Existe_FactD"].ToString() == "SI" ? true : false;

                    if(r["fdesdeCCR"] != System.DBNull.Value)
                        t.fdesdeccr = Convert.ToDateTime(r["fdesdeCCR"]);

                    if (r["fhastaCCR"] != System.DBNull.Value)
                        t.fhastaccr = Convert.ToDateTime(r["fhastaCCR"]);

                    if (r["PotMaxCCR"] != System.DBNull.Value)
                        t.potmaxccr = Convert.ToInt32(r["PotMaxCCR"]);

                    if (r["Maxdefh"] != System.DBNull.Value)
                        t.maxdefh = Convert.ToDateTime(r["Maxdefh"]);

                    //50
                    if (r["Gestion_ATR_Propia"] != System.DBNull.Value)
                       t.gestion_propia_atr = r["Gestion_ATR_Propia"].ToString();

                    if (r["AFactD"] != System.DBNull.Value)
                        t.afactd = Convert.ToDouble(r["AFactD"]);

                    if (r["ExecesosPotencia"] != System.DBNull.Value)
                        t.excesospotencia =Convert.ToDouble(r["ExecesosPotencia"]);

                    if (r["ExcesosReactiva"] != System.DBNull.Value)
                        t.excesosreactiva = Convert.ToDouble(r["ExcesosReactiva"]);

                    if (r["esPrimeraFactura"] != System.DBNull.Value)
                       t.esprimerafactura= r["esPrimeraFactura"].ToString();

                    if (r["aaaamm_ULT_FACT"] != System.DBNull.Value)
                        t.aaaamm_ult_fact = Convert.ToInt32(r["aaaamm_ULT_FACT"]);

                    if (r["fdULTFACT"] != System.DBNull.Value)
                        t.fdULTFACT = Convert.ToDateTime(r["fdULTFACT"]);

                    if (r["fhULTFACT"] != System.DBNull.Value)
                        t.fhULTFACT = Convert.ToDateTime(r["fhULTFACT"]);

                    if (r["A_ULT_FACT"] != System.DBNull.Value)
                        t.a_ULT_FACT = Convert.ToDouble(r["A_ULT_FACT"]);

                    if (r["MinA"] != System.DBNull.Value)
                        t.minA = Convert.ToDouble(r["MinA"]);

                    //60
                    if(r["MaxA"] != System.DBNull.Value)
                        t.maxA = Convert.ToDouble(r["MaxA"]);

                    if (r["MedA"] != System.DBNull.Value)
                        t.medA = Convert.ToDouble(r["MedA"]);

                    if (r["Peor_fuente"] != System.DBNull.Value)
                        t.peor_fuente = r["Peor_fuente"].ToString();

                    if (r["ULT_Fuente"] != System.DBNull.Value)
                        t.ult_fuente = r["ULT_Fuente"].ToString();

                    if (r["TOTAL_DEUDA"] != System.DBNull.Value)
                        t.total_deuda = Convert.ToDouble(r["TOTAL_DEUDA"]);

                    if (r["ult_recep_web"] != System.DBNull.Value)
                        t.ult_recep_web = Convert.ToDateTime(r["ult_recep_web"]);

                    if (r["Grupo"] != System.DBNull.Value)
                        t.grupo = r["Grupo"].ToString();

                    if (r["Ranking_TAM"] != System.DBNull.Value)
                        t.ranking_tam = Convert.ToInt32(r["Ranking_TAM"]);

                    if (r["Ranking_MaxFact"] != System.DBNull.Value)
                        t.ranking_maxfact = Convert.ToInt32(r["Ranking_MaxFact"]);

                    if (r["TAM_por_CUPS"] != System.DBNull.Value)
                        t.tam_por_cups = Convert.ToDouble(r["TAM_por_CUPS"]);

                    //70
                    if (r["factMax"] != System.DBNull.Value)
                        t.fact_max = Convert.ToDouble(r["factMax"]);


                    if (r["diaFactMed"] != System.DBNull.Value)
                        t.diafactmed = Convert.ToInt32(r["diaFactMed"]);

                    if (r["diaFactMax"] != System.DBNull.Value)
                        t.diafacmax = Convert.ToInt32(r["diaFactMax"]);

                    if (r["N_PMs_CS"] != System.DBNull.Value)
                        t.n_pms_cs = r["N_PMs_CS"].ToString();

                    if (r["TLFCS"] != System.DBNull.Value)
                        t.tlfcs = r["TLFCS"].ToString();

                    if (r["TM_OK"] != System.DBNull.Value)
                        t.tm_ok = r["TM_OK"].ToString();

                    if (r["CONSUMO_0"] != System.DBNull.Value)
                        t.consumo_0 = r["CONSUMO_0"].ToString();

                    if (r["PERDIDAS_ML"] != System.DBNull.Value)
                        t.perdidas_ml = r["PERDIDAS_ML"].ToString();

                    if (r["Maxdefh_Absoluta"] != System.DBNull.Value)
                        t.maxdefh_Absoluta = Convert.ToDateTime(r["Maxdefh_Absoluta"]);

                    if (r["AAAAMM_Maxdefh"] != System.DBNull.Value)
                        t.aaaamm_maxdefh = Convert.ToInt32(r["AAAAMM_Maxdefh"]);

                    //80
                    if (r["AAAAMM_Maxdefh_Absoluta"] != System.DBNull.Value)
                        t.aaaamm_maxdefh_absoluta = Convert.ToInt32(r["AAAAMM_Maxdefh_Absoluta"]);

                    if (r["Fecha_Baja"] != System.DBNull.Value)
                        t.fecha_baja = Convert.ToDateTime(r["Fecha_Baja"]);

                    if (r["Proceso_Concursal"] != System.DBNull.Value)
                        t.proceso_concursal = r["Proceso_Concursal"].ToString();

                    if (r["MAXDEFH_CIERRE"] != System.DBNull.Value)
                        t.maxdefh_cierre = Convert.ToDateTime(r["MAXDEFH_CIERRE"]);

                    if (r["AAAAMM_MAXDEFH_CIERRE"] != System.DBNull.Value)
                        t.aaaamm_maxdefh_cierre = Convert.ToInt32(r["AAAAMM_MAXDEFH_CIERRE"]);


                    lista.Add(t);



                }
                db.CloseConnection();
                return lista;
               
            }
            catch (Exception e)
            {
                Console.WriteLine("GeneraInforme: " + e.Message);
                ficheroLog.AddError("GeneraInforme: " + e.Message);
                return null;
            }
        }

        private string GeneraInforme_Estado_Pdte_Origen_SCEA(List<EndesaEntity.medida.SCEA_Tabla> lista)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage;
            FileInfo fileInfo;

            int c = 1;
            int f = 1;

            try
            {
                ficheroLog.Add("Generando Informe");
                ficheroLog.Add("=================");

                fileInfo = new FileInfo(param.GetValue("salida_programada")
                + param.GetValue("prefijo_archivo")
                + DateTime.Now.ToString("yyyy_MM_dd_HHmmss") + ".xlsx");

                excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("SCEA");

                var headerCells = workSheet.Cells[1, 1, 1, 8];
                var headerFont = headerCells.Style.Font;

                var allCells = workSheet.Cells[1, 1, 1, 8];
                var cellFont = allCells.Style.Font;
                cellFont.Bold = true;

                #region Cabecera Informe

                workSheet.Cells[f, c].Value = "IDU";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "CUPS20";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "ET";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "FA";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "FprevB";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Tarifa";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Descripcion_Estado_SCE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "CLIENTE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "NIF";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "c3";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "End";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "tdistri";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "ddistrib";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Tipo_TOP";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Tipo_PM_Calculado";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "PotMaxCont";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "PROVINC";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Coment_SCE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "fComentSCE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Info";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "NOMBRE_GESTOR";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "TERRITORIAL";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "RESPONSABLE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "PMs_PDTES";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "aaaammTrabajo";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "diaHabil";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "PrimerMesPDTE";
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

                workSheet.Cells[f, c].Value = "nMesesPDTES";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Clase";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "usuario";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Grupo_Resolucion";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "aaaammLTP";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Definicion";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "fhLTP";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Anomalia";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "dificultad";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "t";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "ActCCR";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "ReactCCR";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "tipoIncompletitud";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "DescripcionIncompletitud";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Incompletitudes";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Existe_FactD";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "fdesdeCCR";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "fhastaCCR";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "PotMaxCCR";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Maxdefh";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Gestion_ATR_Propia";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "AFactD";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "ExecesosPotencia";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "ExcesosReactiva";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "esPrimeraFactura";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "aaaamm_ULT_FACT";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "fdULTFACT";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "fhULTFACT";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "A_ULT_FACT";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "MinA";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "MaxA";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "MedA";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Peor_fuente";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "ULT_Fuente";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "TOTAL_DEUDA";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "ult_recep_web";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Grupo";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Ranking_TAM";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Ranking_MaxFact";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "TAM_por_CUPS";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "factMax";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "diaFactMed";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "diaFactMax";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "N_PMs_CS";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "TLFCS";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "TM_OK";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "CONSUMO_0";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "PERDIDAS_ML";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Maxdefh_Absoluta";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "AAAAMM_Maxdefh";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "AAAAMM_Maxdefh_Absoluta";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Fecha_Baja";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "Proceso_Concursal";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "MAXDEFH_CIERRE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "AAAAMM_MAXDEFH_CIERRE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                #endregion

                foreach(EndesaEntity.medida.SCEA_Tabla p in lista)
                {
                    c = 1;
                    f++;

                    //1
                    if (p.cup13 != null)
                    {
                        workSheet.Cells[f, c].Value = p.cup13;
                    }
                    c++;

                    if (p.cups20 != null)
                    {
                        workSheet.Cells[f, c].Value = p.cups20;
                    }
                    c++;

                    if (p.empresa != null)
                    {
                        workSheet.Cells[f, c].Value = p.empresa;
                    }
                    c++;

                    if (p.fecha_anexion != DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_anexion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.fecha_prevista_baja != DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_prevista_baja;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.tarifa != null)
                    {
                        workSheet.Cells[f, c].Value = p.tarifa;
                    }
                    c++;

                    if (p.estado_contrato != null)
                    {
                        workSheet.Cells[f, c].Value = p.estado_contrato;
                    }
                    c++;

                    if (p.cliente != null)
                    {
                         workSheet.Cells[f, c].Value = p.cliente;
                    }
                    c++;

                    if (p.nif != null)
                    {
                        workSheet.Cells[f, c].Value = p.nif;
                    }
                    c++;

                    //10
                    if (p.c3 != null)
                    {
                        workSheet.Cells[f, c].Value = p.c3;
                    }
                    c++;

                    if (p.end != null)
                    {
                        workSheet.Cells[f, c].Value = p.end;
                    }
                    c++;

                    if (p.tipo_distribora != null)
                    {
                        workSheet.Cells[f, c].Value = p.tipo_distribora;
                    }
                    c++;

                    if (p.descripcion_distribuidora != null)
                    {
                        workSheet.Cells[f, c].Value = p.descripcion_distribuidora;
                    }
                    c++;

                    if (p.tipo_top != null)
                    {
                        workSheet.Cells[f, c].Value = p.tipo_top;
                    }
                    c++;

                    if (p.tipo_pm_calculado != 0)
                    {
                        workSheet.Cells[f, c].Value = p.tipo_pm_calculado;
                    }
                    c++;

                    if (p.potencia_maxima_contratada != 0)
                    {
                        workSheet.Cells[f, c].Value = p.potencia_maxima_contratada;
                    }
                    c++;

                    if (p.provincia != null)
                    {
                        workSheet.Cells[f, c].Value = p.provincia;
                    }
                    c++;

                    if (p.comentario_sce != null)
                    {
                        workSheet.Cells[f, c].Value = p.comentario_sce;
                    }
                    c++;

                    if (p.fecha_comentario_sce != null)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_comentario_sce;
                    }
                    c++;

                    //20
                    if (p.info != null)
                    {
                        workSheet.Cells[f, c].Value = p.info;
                    }
                    c++;

                    if (p.nombre_gestor != null)
                    {
                        workSheet.Cells[f, c].Value = p.nombre_gestor;
                    }
                    c++;

                    if (p.territorial != null)
                    {
                        workSheet.Cells[f, c].Value = p.territorial;
                    }
                    c++;

                    if (p.responsable != null)
                    {
                        workSheet.Cells[f, c].Value = p.responsable;
                    }
                    c++;

                    if (p.pm_pdtes != 0)
                    {
                        workSheet.Cells[f, c].Value = p.pm_pdtes;
                    }
                    c++;
                    if (p.aaaamm_trabajo != 0)
                    {
                        workSheet.Cells[f, c].Value = p.aaaamm_trabajo;
                    }
                    c++;
                    if (p.dia_habil != 0)
                    {
                        workSheet.Cells[f, c].Value = p.dia_habil;
                    }
                    c++;

                    if (p.primer_mes_pdte != 0)
                    {
                        workSheet.Cells[f, c].Value = p.primer_mes_pdte;
                    }
                    c++;

                    if (p.estado != null)
                    {
                        workSheet.Cells[f, c].Value = p.estado;
                    }
                    c++;

                    if (p.subestado != null)
                    {
                        workSheet.Cells[f, c].Value = p.subestado;
                    }
                    c++;

                    //30
                    if (p.num_meses_pdtes != 0)
                    {
                        workSheet.Cells[f, c].Value = p.num_meses_pdtes;
                    }
                    c++;

                    if (p.clase != 0)
                    {
                        workSheet.Cells[f, c].Value = p.clase;
                    }
                    c++;

                    if (p.usuario != null)
                    {
                        workSheet.Cells[f, c].Value = p.usuario;
                    }
                    c++;

                    if (p.grupo_resolucion != null)
                    {
                        workSheet.Cells[f, c].Value = p.grupo_resolucion;
                    }
                    c++;

                    if (p.aaaammlpt != 0)
                    {
                        workSheet.Cells[f, c].Value = p.aaaammlpt;
                    }
                    c++;

                    if (p.definicion != null)
                    {
                        workSheet.Cells[f, c].Value = p.definicion;
                    }
                    c++;

                    if (p.fechahasta_ltp != DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fechahasta_ltp;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.anomalia != null)
                    {
                        workSheet.Cells[f, c].Value = p.anomalia;
                    }
                    c++;

                    if (p.dificultad != 0)
                    {
                        workSheet.Cells[f, c].Value = p.dificultad;
                    }
                    c++;

                    if (p.t != null)
                    {
                        workSheet.Cells[f, c].Value = p.t;
                    }
                    c++;

                    //40
                    if (p.activa_ccr != 0)
                    {
                        workSheet.Cells[f, c].Value = p.activa_ccr;
                    }
                    c++;

                    if (p.reactiva_ccr != 0)
                    {
                        workSheet.Cells[f, c].Value = p.reactiva_ccr;
                    }
                    c++;

                    if (p.tipo_incompletitud != null)
                    {
                        workSheet.Cells[f, c].Value = p.tipo_incompletitud;
                    }
                    c++;

                    if (p.descripcion_incompletitud != null)
                    {
                        workSheet.Cells[f, c].Value =p.descripcion_incompletitud;
                    }
                    c++;

                    if (p.incompletitudes != 0)
                    {
                        workSheet.Cells[f, c].Value = p.incompletitudes;
                    }
                    c++;
                                        
                    workSheet.Cells[f, c].Value = p.existe_FactD ? "SI" : "NO";
                    c++;

                    if (p.fdesdeccr != DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fdesdeccr;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.fhastaccr != DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fhastaccr;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.potmaxccr != 0)
                    {
                        workSheet.Cells[f, c].Value = p.potmaxccr;
                    }
                    c++;

                    if (p.maxdefh != DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.maxdefh;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    //50
                    if (p.gestion_propia_atr != null)
                    {
                        workSheet.Cells[f, c].Value = p.gestion_propia_atr;
                    }
                    c++;

                    if (p.afactd != 0)
                    {
                        workSheet.Cells[f, c].Value = p.afactd;
                    }
                    c++;

                    if (p.excesospotencia != 0)
                    {
                        workSheet.Cells[f, c].Value = p.excesospotencia;
                    }
                    c++;

                    if (p.excesosreactiva != 0)
                    {
                        workSheet.Cells[f, c].Value = p.excesosreactiva;
                    }
                    c++;

                    if (p.esprimerafactura != null)
                    {
                        workSheet.Cells[f, c].Value = p.esprimerafactura;
                    }
                    c++;

                    if (p.aaaamm_ult_fact != 0)
                    {
                        workSheet.Cells[f, c].Value = p.aaaamm_ult_fact;
                    }
                    c++;

                    if (p.fdULTFACT != DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fdULTFACT;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.fhULTFACT != DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fhULTFACT;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.a_ULT_FACT != 0)
                    {
                        workSheet.Cells[f, c].Value = p.a_ULT_FACT;
                    }
                    c++;

                    if (p.minA != 0)
                    {
                        workSheet.Cells[f, c].Value = p.minA;
                    }
                    c++;

                    //60
                    if (p.maxA != 0)
                    {
                        workSheet.Cells[f, c].Value = p.maxA;
                    }
                    c++;

                    if (p.medA != 0)
                    {
                        workSheet.Cells[f, c].Value = p.medA;
                    }
                    c++;

                    if (p.peor_fuente != null)
                    {
                        workSheet.Cells[f, c].Value = p.peor_fuente;
                    }
                    c++;

                    if (p.ult_fuente != null)
                    {
                        workSheet.Cells[f, c].Value = p.ult_fuente;
                    }
                    c++;

                    if (p.total_deuda != 0)
                    {
                        workSheet.Cells[f, c].Value = p.total_deuda;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    }
                    c++;

                    if (p.ult_recep_web != DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.ult_recep_web;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.grupo != null)
                    {
                        workSheet.Cells[f, c].Value = p.grupo;
                    }
                    c++;

                    if (p.ranking_tam != 0)
                    {
                        workSheet.Cells[f, c].Value = p.ranking_tam;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (p.ranking_maxfact != 0)
                    {
                        workSheet.Cells[f, c].Value = p.ranking_maxfact;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    }
                    c++;

                    if (p.tam_por_cups != 0)
                    {
                        workSheet.Cells[f, c].Value = p.tam_por_cups;
                    }
                    c++;


                    //70
                    if (p.fact_max != 0)
                    {
                        workSheet.Cells[f, c].Value = p.fact_max;
                    }
                    c++;

                    if (p.diafactmed != 0)
                    {
                        workSheet.Cells[f, c].Value = p.diafactmed;
                    }
                    c++;

                    if (p.diafacmax != 0)
                    {
                        workSheet.Cells[f, c].Value = p.diafacmax;
                    }
                    c++;

                    if (p.n_pms_cs != null)
                    {
                        workSheet.Cells[f, c].Value = p.n_pms_cs;
                    }
                    c++;

                    if (p.tlfcs != null)
                    {
                        workSheet.Cells[f, c].Value = p.tlfcs;
                    }
                    c++;

                    if (p.tm_ok != null)
                    {
                        workSheet.Cells[f, c].Value = p.tm_ok;
                    }
                    c++;

                    if (p.consumo_0 != null)
                    {
                        workSheet.Cells[f, c].Value = p.consumo_0;
                    }
                    c++;

                    if (p.perdidas_ml != null)
                    {
                        workSheet.Cells[f, c].Value = p.perdidas_ml;
                    }
                    c++;

                    if (p.maxdefh_Absoluta != DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.maxdefh_Absoluta;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.aaaamm_maxdefh != 0)
                    {
                        workSheet.Cells[f, c].Value = p.aaaamm_maxdefh;
                    }
                    c++;

                    //80
                    if (p.aaaamm_maxdefh_absoluta != 0)
                    {
                        workSheet.Cells[f, c].Value = p.aaaamm_maxdefh_absoluta;
                    }
                    c++;

                    if (p.fecha_baja != DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_baja;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.proceso_concursal != null)
                    {
                        workSheet.Cells[f, c].Value = p.proceso_concursal;
                    }
                    c++;

                    if (p.maxdefh_cierre != DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.maxdefh_cierre;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.aaaamm_maxdefh_cierre != 0)
                    {
                        workSheet.Cells[f, c].Value = p.aaaamm_maxdefh_cierre;
                    }
                    
                }

                headerCells = workSheet.Cells[1, 1, 1, c];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;
                allCells = workSheet.Cells[1, 1, f, c];

                allCells.AutoFitColumns();

                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:CF1"].AutoFilter = true;
                allCells.AutoFitColumns();

                excelPackage.Save();
                Console.WriteLine("Excel Generado");

                return fileInfo.FullName;
            }
            catch(Exception e)
            {
                return null;
            }
        }

        private void EnvioCorreo(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                string from = param.GetValue("mail_from");
                string to = param.GetValue("mail_to");
                string cc = param.GetValue("mail_cc");
                string subject = param.GetValue("mail_subject") + " " + DateTime.Now.ToString("dd/MM/yyyy");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Se adjunta el archivo ").Append(fileInfo.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");

                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (param.GetValue("mail_enviar") == "S")
                    mes.SendMail(from, to, cc, subject, textBody.ToString(), archivo);

                else
                    mes.SaveMail(from, to, cc, subject, textBody.ToString(), archivo);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("mail_from"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }
    }
}
