using EndesaBusiness.contratacion;
using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using EndesaEntity.cnmc.V21_2019_12_17;
using EndesaEntity.facturacion;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Graph;
using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion.redshift
{
       
    public class Pendiente : EndesaEntity.medida.Pendiente
    {
        utilidades.Param param;
        utilidades.Seguimiento_Procesos ss_pp;
        logs.Log ficheroLog;

        Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> dic_pendiente_hist_fecha;
        Dictionary<string, EndesaEntity.medida.Pendiente> dic_pendiente;
        Dictionary<string, DateTime> dic_dias_estado;

        public Pendiente()
        {
            param = new utilidades.Param("t_ed_h_sap_pendiente_param", servidores.MySQLDB.Esquemas.FAC);
            ss_pp = new utilidades.Seguimiento_Procesos();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Copia_Pendiente_BI");
            dic_pendiente = Carga();
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
                    + " pend.cups13, "
                    + " pend.mes as aaaammPdte, pend.estado, pend.subestado"
                    + " FROM fact.t_ed_h_sap_pendiente_facturar pend"
                    + " ORDER BY pend.empresa_titular, "
                    + " pend.cups13, pend.mes ASC";

                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
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

        public void CopiaDatos()
        {
            StringBuilder sb = new StringBuilder();
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";

            MySQLDB dbmy;
            MySqlCommand commandmy;

            bool firstOnly = true;
            int j = 0;
            int k = 0;
            int totalRegistros = 0;

            DateTime ultimaFechaCopiado = new DateTime();
            DateTime fechaInformeBI = new DateTime();

            utilidades.Fechas utilfecha = new Fechas();

            try
            {                

                ultimaFechaCopiado = UltimaActualizacionMySQL();
                //fechaInformeBI = UltimaActualizacionBI();

                if (ultimaFechaCopiado < DateTime.Now.Date)
                {
                    ss_pp.Update_Fecha_Inicio("Facturación", "Copia Pendiente BI", "Copia Pendiente BI");

                    //borrado_tabla();                    

                    ficheroLog.Add(Consulta(ultimaFechaCopiado));
                    db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                    command = new OdbcCommand(Consulta(ultimaFechaCopiado.AddMonths(-1)), db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        j++;
                        k++;

                        if (firstOnly)
                        {
                            sb = null;
                            sb = new StringBuilder();
                            sb.Append("REPLACE INTO t_ed_h_sap_pendiente_facturar");
                            sb.Append(" (cd_cups, id_instalacion, cl_stro, id_crto_ext, cl_crto_ext, cd_empr_distdora, fh_desde, fh_hasta,");
                            sb.Append("fh_periodo, cd_estado, cd_subestado, lg_multimedida, cd_empr_titular, cd_ritmo_fact,");
                            sb.Append("cd_segmento_ptg, fh_envio, fec_act, cod_carga, agora, tam) values ");

                            firstOnly = false;
                        }
                        #region campos

                        if (r["cd_cups"] != System.DBNull.Value)
                            sb.Append("('").Append(r["cd_cups"].ToString()).Append("',");
                        else
                            sb.Append("(null,");

                        if (r["id_instalacion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["id_instalacion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cl_stro"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cl_stro"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["id_crto_ext"] != System.DBNull.Value)
                            sb.Append("'").Append(r["id_crto_ext"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cl_crto_ext"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cl_crto_ext"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_empr_distdora"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_empr_distdora"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_desde"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_desde"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_hasta"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_hasta"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_periodo"] != System.DBNull.Value)
                            sb.Append(r["fh_periodo"].ToString()).Append(",");
                        else
                            sb.Append("null,");

                        if (r["cd_estado"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_estado"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_subestado"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_subestado"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_multimedida"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_multimedida"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_empr_titular"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_empr_titular"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_ritmo_fact"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_ritmo_fact"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_segmento_ptg"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_segmento_ptg"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_envio"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_envio"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fec_act"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fec_act"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cod_carga"] != System.DBNull.Value)
                            sb.Append(r["cod_carga"].ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        //AGORA                        
                        if (r["lg_agora"] != System.DBNull.Value)
                            sb.Append("'S',");
                        else
                            sb.Append("'N',");

                        if (r["nm_tam"] != System.DBNull.Value)
                            sb.Append(r["nm_tam"].ToString().Replace(",", ".")).Append("),");
                        else
                            sb.Append("null),");


                        #endregion



                        if (j == 100)
                        {
                            Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                            firstOnly = true;
                            dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                            commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                            commandmy.ExecuteNonQuery();
                            dbmy.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            j = 0;
                        }


                    }
                    db.CloseConnection();

                    if (j > 0)
                    {
                        Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                        firstOnly = true;
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                        commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        j = 0;
                    }

                    Construccion_Datos();

                    ss_pp.Update_Fecha_Fin("Facturación", "Copia Pendiente BI", "Copia Pendiente BI");

                }
                
                

            }
            catch(Exception ex)
            {
                ficheroLog.addError(ex.Message);
            }



        }

        private void borrado_tabla()
        {
            MySQLDB db;
            MySqlCommand command;
            
            string strSql = "";

            strSql = "delete from t_ed_h_sap_pendiente_facturar";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        private DateTime UltimaActualizacionMySQL()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime fecha = new DateTime(2022, 01, 01);

            strSql = "SELECT max(fh_envio) AS fh_envio FROM t_ed_h_sap_pendiente_facturar";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
                if (r["fh_envio"] != System.DBNull.Value)
                    fecha = Convert.ToDateTime(r["fh_envio"]);
            db.CloseConnection();
            Console.WriteLine("Última fecha de copiado MySQL: " + fecha.ToString("dd/MM/yyyy"));
            ficheroLog.Add("Última fecha de copiado MySQL: " + fecha.ToString("dd/MM/yyyy"));

            return fecha;
        }

        private DateTime UltimaActualizacionBI()
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";
            DateTime fecha = new DateTime(2022, 01, 01);

            strSql = "SELECT max(fh_envio) AS fh_envio FROM ed_owner.t_ed_h_sap_pendiente_facturar";
            db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
            command = new OdbcCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())                
                if (r["fh_envio"] != System.DBNull.Value)
                    fecha = Convert.ToDateTime(r["fh_envio"]);
            db.CloseConnection();
            Console.WriteLine("Última fecha de copiado BI: " + fecha.ToString("dd/MM/yyyy"));
            ficheroLog.Add("Última fecha de copiado BI: " + fecha.ToString("dd/MM/yyyy"));

            return fecha.AddDays(-4);
        }

        private string Consulta(DateTime f)
        {
            string strSql = "";

            strSql = "SELECT p.cd_cups, i.id_instalacion, p.cl_stro, p.id_crto_ext, p.cl_crto_ext, p.cd_empr_distdora,"
                + " p.fh_desde, p.fh_hasta, p.fh_periodo, p.cd_estado, p.cd_subestado, p.lg_multimedida,"
                + " p.cd_empr_titular, p.cd_ritmo_fact, p.cd_segmento_ptg, p.fh_envio, p.fec_act, p.cod_carga,"
                + " i.nm_tam, i.lg_agora"
                + " FROM ed_owner.t_ed_h_sap_pendiente_facturar p"
                + " left outer join ed_owner.t_ed_h_sap_instalacion i on"
                + " i.cd_cups = p.cd_cups"
                + " where"
                + " fh_envio > '" + f.ToString("yyyy-MM-dd") + "'";

            return strSql;

        }

        public void Construccion_Datos()
        {
            MySQLDB db;
            MySqlCommand command;            
            string strSql = "";

            strSql = "REPLACE INTO  t_ed_h_sap_pendiente_facturar_agrupado"
                + " SELECT substr(cd_cups,1,20) as cd_cups, id_instalacion, cl_stro, id_crto_ext, cl_crto_ext,"
                + " cd_empr_distdora, fh_desde, fh_hasta, fh_periodo, cd_estado, cd_subestado,"
                + " lg_multimedida, cd_empr_titular, cd_ritmo_fact, cd_segmento_ptg, fh_envio,"
                + " fec_act, cod_carga, TAM, agora, now()"
                + " FROM t_ed_h_sap_pendiente_facturar p"
                + " ORDER BY cd_cups, fh_periodo DESC";
            ficheroLog.Add(strSql);
            Console.WriteLine(strSql);
            db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "replace into t_ed_h_sap_tam_agora"
            + " SELECT g.cd_cups, g.id_instalacion, g.id_crto_ext,"
            + " g.cd_empr_titular, g.cd_segmento_ptg, g.tam, g.agora, g.fec_act"
            + " FROM t_ed_h_sap_pendiente_facturar_agrupado g"
            + " ORDER BY g.fec_act";
            ficheroLog.Add(strSql);
            Console.WriteLine(strSql);
            db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


        }

        public void GeneraInformePendSAP(bool automatico)
        {
            FileInfo file;
            string ruta_salida_archivo = "";            

            string[] listaArchivos = System.IO.Directory.GetFiles(automatico ? param.GetValue("ruta_salida_informe") : @"c:\Temp\",
                    param.GetValue("prefijo_informe") + "*.xlsx");

            for (int i = 0; i < listaArchivos.Length; i++)
            {
                file = new FileInfo(listaArchivos[i]);
                file.Delete();
            }

            if(automatico)
                ruta_salida_archivo = param.GetValue("ruta_salida_informe")
                    + param.GetValue("prefijo_informe")
                    + DateTime.Now.ToString("yyyyMMdd")
                    + param.GetValue("sufijo_informe");
            else
                ruta_salida_archivo = @"c:\Temp\"
                   + param.GetValue("prefijo_informe")
                   + DateTime.Now.ToString("yyyyMMdd")
                   + param.GetValue("sufijo_informe");

            InformePendiente_BI_Facturacion(ruta_salida_archivo, automatico);
        }

        private void InformePendiente_BI_Facturacion(string ruta_salida_archivo, bool automatico)
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

            bool tiene_complemento_a01 = false;
            bool sacar_portugal = true;


            DateTime fd = new DateTime();
            DateTime fd_tam = new DateTime();
            DateTime udh = new DateTime();
                       


            utilidades.Fechas utilfecha = new Fechas();

            try
            {                

                if (!automatico || (UltimaActualizacionMySQL().Date > 
                    ss_pp.GetFecha_FinProceso("Facturación", "Informe Pendiente BI", "Informe Pendiente BI").Date))
                {

                    if(automatico)
                        ss_pp.Update_Fecha_Inicio("Facturación", "Informe Pendiente BI", "Informe Pendiente BI");


                    FileInfo plantillaExcel =
                        new FileInfo(System.Environment.CurrentDirectory +
                        param.GetValue("plantilla_informe_pendiente"));

                    FileInfo fileInfo = new FileInfo(ruta_salida_archivo);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                    ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);

                    var workSheet = excelPackage.Workbook.Worksheets["Resumen ES"];
                    var headerCells = workSheet.Cells[1, 1, 1, 17];
                    var headerFont = headerCells.Style.Font;

                    List<string> lista_empresas_ES = new List<string>();
                    lista_empresas_ES.Add("ES21");
                    lista_empresas_ES.Add("ES22");

                    List<string> lista_empresas_PT = new List<string>();
                    lista_empresas_PT.Add("PT1Q");

                    List<string> lista_segmentos_MT_BTE = new List<string>();
                    lista_segmentos_MT_BTE.Add("MT");
                    lista_segmentos_MT_BTE.Add("BTE");

                    List<string> lista_segmentos_BTN = new List<string>();
                    lista_segmentos_BTN.Add("BTN");

                    // Tomamos lo últimos 5 días hábiles
                    // Si se lanza el listado en día fin de semana
                    // hay que quitar un día más porque el viernes todavía no se ha procesado



                    if (!utilfecha.EsLaborable())
                    {
                        fd = utilfecha.UltimoDiaHabilAnterior(
                               utilfecha.UltimoDiaHabilAnterior(
                                   utilfecha.UltimoDiaHabilAnterior(
                                       utilfecha.UltimoDiaHabilAnterior(
                                           utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil())))));

                        udh = utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil());
                    }
                    else
                    {
                        fd = utilfecha.UltimoDiaHabilAnterior(
                                utilfecha.UltimoDiaHabilAnterior(
                                    utilfecha.UltimoDiaHabilAnterior(
                                        utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()))));

                        udh = utilfecha.UltimoDiaHabil();
                    }

                    

                    fd_tam = utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil());


                   


                    int totales_dia = 0;
                    double totales_dia_tam = 0;

                    bool noAgora = false;
                    bool siAgora = true;

                    

                    



                    Dictionary<DateTime, int> dic_Totales_cups = new Dictionary<DateTime, int>();
                    Dictionary<DateTime, double> dic_Totales_tam = new Dictionary<DateTime, double>();


                    dic_dias_estado = CargaDiasEstado();


                    int dia = 0;                  

                    //  ES

                    dic_pendiente_hist_fecha = CargaPendienteHist_DesdeFecha(CalculaFechaDesdeinforme(), lista_empresas_ES);

                    f = 4;
                    c = 9;
                    dic_Totales_cups = GeneraResumen(dic_Totales_cups, excelPackage, "Resumen ES", true, noAgora, lista_empresas_ES, null,
                        dic_pendiente_hist_fecha, dia, f, c, false);


                    f = 58;
                    c = 9;
                    dic_Totales_cups = GeneraResumen(dic_Totales_cups, excelPackage, "Resumen ES", false, siAgora, lista_empresas_ES, null,
                        dic_pendiente_hist_fecha, dia, f, c, true);

                    #region NO ÁGORA ES

                    //foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                    //{


                    //    dia++;

                    //    if (dia < 6)
                    //    {

                    //        Console.WriteLine("Totales ES noAgora dia: " + p.Key.ToString("dd/MM/yyyy"));

                    //        f = 4;
                    //        //c++;
                    //        c--;

                    //        workSheet.Cells[f, c].Value = p.Key;
                    //        workSheet.Cells[f, c].Style.Font.Bold = true;
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    //        f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", "01.A");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", "01.B01");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", "01.B02");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", "01.B03");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", "01.C");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", "01.D");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;                            
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", "01.F");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", "01.G");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", "01.H");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", "01.I");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", "01.J");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", "01.K");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", "01.Z");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    //        totales_noagora_01 =
                    //              Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "01", null);

                    //        workSheet.Cells[f, c].Value = totales_noagora_01;
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //        workSheet.Cells[f, c].Style.Font.Bold = true;
                    //        f++;

                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "02", "02.A");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "02", "02.B01");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;                            
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "02", "02.B05");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "02", "02.B06");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    //        totales_noagora_02 =
                    //              Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "02", null);

                    //        workSheet.Cells[f, c].Value = totales_noagora_02;
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //        workSheet.Cells[f, c].Style.Font.Bold = true;
                    //        f++;

                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null, "03", "03.B01");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null, "03", "03.B02");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null, "03", "03.B03");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null, "03", "03.B0Z");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null, "03", "03.B11");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null, "03", "03.B12");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null, "03", "03.B15");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "03", "03.C");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "03", "03.D");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "03", "03.F");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "03", "03.G");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "03", "03.H");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "03", "03.I");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "03", "03.J");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "03", "03.L");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "03", "03.Z");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    //        totales_noagora_03 =
                    //              Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "03", null);

                    //        workSheet.Cells[f, c].Value = totales_noagora_03;
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //        workSheet.Cells[f, c].Style.Font.Bold = true;
                    //        f++;

                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "04", "04.A");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "04", "04.B01");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "04", "04.B02");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "04", "04.B03");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "04", "04.B04");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "04", "04.B05");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "04", "04.B0Z");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null, "04", "04.B11");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null, "04", "04.B12");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null, "04", "04.B13");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null, "04", "04.B14");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null, "04", "04.B15");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    //        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "04", "04.Z");
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    //        totales_noagora_04 =
                    //             Total_Pendiente(noAgora, p.Key, lista_empresas_ES, null,  "04", null);

                    //        workSheet.Cells[f, c].Value = totales_noagora_04;
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //        workSheet.Cells[f, c].Style.Font.Bold = true;
                    //        f++;


                    //        workSheet.Cells[f, c].Value = totales_noagora_01
                    //            + totales_noagora_02
                    //            + totales_noagora_03
                    //            + totales_noagora_04;
                    //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //        workSheet.Cells[f, c].Style.Font.Bold = true;
                    //        f++;


                    //        int o;
                    //        if (!dic_Totales_cups.TryGetValue(p.Key, out o))
                    //            dic_Totales_cups.Add(p.Key, totales_noagora_01
                    //            + totales_noagora_02
                    //            + totales_noagora_03
                    //            + totales_noagora_04);

                    //    }
                    //}

                    #endregion


                    f = 4;
                    c = 14;
                    dic_Totales_tam = GeneraResumenTAM(dic_Totales_tam, excelPackage, "Resumen ES", true, noAgora, lista_empresas_ES, null,
                        dic_pendiente_hist_fecha, dia, f, c, false);

                    f = 58;
                    c = 14;
                    dic_Totales_tam = GeneraResumenTAM(dic_Totales_tam, excelPackage, "Resumen ES", false, siAgora, lista_empresas_ES, null,
                        dic_pendiente_hist_fecha, dia, f, c, true);


                    dic_Totales_cups.Clear();
                    dic_Totales_tam.Clear();


                    // PT MT BTE

                    dic_pendiente_hist_fecha = CargaPendienteHist_PT_DesdeFecha(CalculaFechaDesdeinforme(), lista_empresas_PT, lista_segmentos_MT_BTE);

                    f = 4;
                    c = 9;
                    dic_Totales_cups = GeneraResumen(dic_Totales_cups, excelPackage, "Resumen POR MT-BTE", true, noAgora, lista_empresas_PT, lista_segmentos_MT_BTE,
                        dic_pendiente_hist_fecha, dia, f, c, false);


                    f = 58;
                    c = 9;
                    dic_Totales_cups = GeneraResumen(dic_Totales_cups, excelPackage, "Resumen POR MT-BTE", false, siAgora, lista_empresas_PT, lista_segmentos_MT_BTE,
                        dic_pendiente_hist_fecha, dia, f, c, true);

                    
                    f = 4;
                    c = 14;
                    dic_Totales_tam = GeneraResumenTAM(dic_Totales_tam, excelPackage, "Resumen POR MT-BTE", true, noAgora, lista_empresas_PT, lista_segmentos_MT_BTE,
                        dic_pendiente_hist_fecha, dia, f, c, false);

                    f = 58;
                    c = 14;
                    dic_Totales_tam = GeneraResumenTAM(dic_Totales_tam, excelPackage, "Resumen POR MT-BTE", false, siAgora, lista_empresas_PT, lista_segmentos_MT_BTE,
                        dic_pendiente_hist_fecha, dia, f, c, true);


                    dic_Totales_cups.Clear();
                    dic_Totales_tam.Clear();


                    // PT BTN

                    dic_pendiente_hist_fecha = CargaPendienteHist_PT_DesdeFecha(CalculaFechaDesdeinforme(), lista_empresas_PT, lista_segmentos_BTN);

                    f = 4;
                    c = 9;
                    dic_Totales_cups = GeneraResumen(dic_Totales_cups, excelPackage, "Resumen POR BTN", true, noAgora, lista_empresas_PT, lista_segmentos_BTN,
                        dic_pendiente_hist_fecha, dia, f, c, false);


                    f = 58;
                    c = 9;
                    dic_Totales_cups = GeneraResumen(dic_Totales_cups, excelPackage, "Resumen POR BTN", false, siAgora, lista_empresas_PT, lista_segmentos_BTN,
                        dic_pendiente_hist_fecha, dia, f, c, true);


                    f = 4;
                    c = 14;
                    dic_Totales_tam = GeneraResumenTAM(dic_Totales_tam, excelPackage, "Resumen POR BTN", true, noAgora, lista_empresas_PT, lista_segmentos_BTN,
                        dic_pendiente_hist_fecha, dia, f, c, false);

                    f = 58;
                    c = 14;
                    dic_Totales_tam = GeneraResumenTAM(dic_Totales_tam, excelPackage, "Resumen POR BTN", false, siAgora, lista_empresas_PT, lista_segmentos_BTN,
                        dic_pendiente_hist_fecha, dia, f, c, true);


                    dic_Totales_cups.Clear();
                    dic_Totales_tam.Clear();


                    #region Detalle ES
                                       

                    DateTime fecha_informe = CalculaFechaDetalle();
                    
                    GeneraDetalle(DetalleExcel(fecha_informe, lista_empresas_ES, null), excelPackage, "Detalle ES", false);
                    GeneraDetalle(DetalleExcel(fecha_informe, lista_empresas_PT, lista_segmentos_MT_BTE), excelPackage, "Detalle POR MT-BTE", true);
                    GeneraDetalle(DetalleExcel(fecha_informe, lista_empresas_PT, lista_segmentos_BTN), excelPackage, "Detalle POR BTN", true);

                    #endregion




                    excelPackage.SaveAs(fileInfo);


                    

                    if (automatico && param.GetValue("mail_enviar_mail_psat_tam") == "S")
                    {
                        ss_pp.Update_Fecha_Fin("Facturación", "Informe Pendiente BI", "Informe Pendiente BI");
                        EnvioCorreo_PdteWeb_BI(ruta_salida_archivo, fecha_informe);
                    }

                    if(automatico)
                        ss_pp.Update_Fecha_Fin("Facturación", "Informe Pendiente BI", "Informe Pendiente BI");

                }
                else if(automatico)
                {
                    ss_pp.Update_Comentario("Facturación", "Informe Pendiente BI", "Informe Pendiente BI",
                        "La fecha de actualización en BI es: " + UltimaActualizacionMySQL().Date.ToString("dd/MM/yyyy"));
                }
            }
        
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

            }
        }



        private Dictionary<DateTime, int> GeneraResumen(Dictionary<DateTime, int> dic_Totales_cups, ExcelPackage excelPackage, string hoja, bool pintar_fechas, bool agora, List<string> lista_empresas,
            List<string> lista_segmentos, Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> dic, int dia, int fila, int c, bool pintar_total)
        {
            int f;

            var workSheet = excelPackage.Workbook.Worksheets[hoja];
            var headerCells = workSheet.Cells[1, 1, 1, 17];
            var headerFont = headerCells.Style.Font;
            

            foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
            {               

                dia++;
                f = fila;

                if (dia < 6)
                {

                    Console.WriteLine("Totales ES noAgora dia: " + p.Key.ToString("dd/MM/yyyy"));
                                        
                    c--;

                    if (pintar_fechas)
                    {
                        workSheet.Cells[f, c].Value = p.Key;
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        
                    }

                    f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.A");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.B01");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.B02");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.B03");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.C");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.D");                    
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.G");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.H");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.I");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.J");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.K");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.Z");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                   

                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", null);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "02", "02.A");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "02", "02.B01");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "02", "02.B05");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "02", "02.B06");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;                    

                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "02", null);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B01");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B02");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B03");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B0Z");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B11");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B12");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B15");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.C");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.D");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.F");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.G");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.H");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.I");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.J");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.L");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.Z");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                                       

                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", null);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.A");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B01");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B02");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B03");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B04");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B05");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B0Z");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B11");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B12");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B13");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B14");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B15");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.Z");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                                      

                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", null);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;


                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "05", "05.A");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "05", "05.B");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "05", "05.C");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "05", null);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;


                    workSheet.Cells[f, c].Value = Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", null)
                        + Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "02", null)
                        + Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", null)
                        + Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", null)
                        + Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "05", null);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;


                    int o;
                    if (!dic_Totales_cups.TryGetValue(p.Key, out o))
                        dic_Totales_cups.Add(p.Key, Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", null)
                        + Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "02", null)
                        + Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", null)
                        + Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", null)
                        + Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "05", null));
                    else
                        o = o + Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "01", null)
                        + Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "02", null)
                        + Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "03", null)
                        + Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "04", null)
                        + Total_Pendiente(agora, p.Key, lista_empresas, lista_segmentos, "05", null);


                    if (pintar_total)
                    {
                        workSheet.Cells[f, c].Value = o;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                    }

                }

                

            }

            return dic_Totales_cups;
        }

        private Dictionary<DateTime, double> GeneraResumenTAM(Dictionary<DateTime, double> dic_Totales_tam, ExcelPackage excelPackage, string hoja, bool pintar_fechas, bool agora, List<string> lista_empresas,
            List<string> lista_segmentos, Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> dic, int dia, int fila, int c, bool pintar_total)
        {
            int f;

            var workSheet = excelPackage.Workbook.Worksheets[hoja];
            var headerCells = workSheet.Cells[1, 1, 1, 17];
            var headerFont = headerCells.Style.Font;

            

            foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
            {


                dia++;
                f = fila;

                if (dia < 6)
                {

                    Console.WriteLine("Totales ES noAgora dia: " + p.Key.ToString("dd/MM/yyyy"));

                    c--;

                    if (pintar_fechas)
                    {
                        workSheet.Cells[f, c].Value = p.Key;
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                    }

                    f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.A");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.B01");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.B02");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.B03");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.C");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.D");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;                    
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.G");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.H");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.I");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.J");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.K");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", "01.Z");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                                       

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", null);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "02", "02.A");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "02", "02.B01");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;                                        
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "02", "02.B05");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "02", "02.B06");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                                      

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "02", null);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;
                    
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B01");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B02");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B04");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B0Z");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B11");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B12");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.B15");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.C");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.D");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.F");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.G");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.H");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.I");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.J");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.L");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", "03.Z");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                                        

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", null);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.A");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B01");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B02");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B03");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B04");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B05");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B0Z");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B11");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B12");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B13");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B14");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.B15");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", "04.Z");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                    

                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", null);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;


                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "05", "05.A");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "05", "05.B");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "05", "05.C");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;


                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "05", null);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;


                    workSheet.Cells[f, c].Value = Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", null)
                        + Total_Pendiente_TAM(agora, p.Key, lista_empresas, null, "02", null)
                        + Total_Pendiente_TAM(agora, p.Key, lista_empresas, null, "03", null)
                        + Total_Pendiente_TAM(agora, p.Key, lista_empresas, null, "04", null)
                        + Total_Pendiente_TAM(agora, p.Key, lista_empresas, null, "05", null);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    f++;


                    double o;
                    if (!dic_Totales_tam.TryGetValue(p.Key, out o))
                        dic_Totales_tam.Add(p.Key, Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", null)
                        + Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "02", null)
                        + Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", null)
                        + Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", null)
                        + Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "05", null));
                    else
                        o = o + Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "01", null)
                        + Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "02", null)
                        + Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "03", null)
                        + Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "04", null)
                        + Total_Pendiente_TAM(agora, p.Key, lista_empresas, lista_segmentos, "05", null);

                    if (pintar_total)
                    {
                        workSheet.Cells[f, c].Value = o;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                    }

                }
            }

            return dic_Totales_tam;
        }

        private void GeneraDetalle(string strSql, ExcelPackage excelPackage, string hoja, bool portugal)
        {
            int f = 1;
            int c = 1;

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;            

            var workSheet = excelPackage.Workbook.Worksheets[hoja];
            var headerCells = workSheet.Cells[1, 1, 1, 30];
            var headerFont = headerCells.Style.Font;            
            headerFont.Bold = true;
            var allCells = workSheet.Cells[1, 1, 50, 50];

            workSheet.View.FreezePanes(2, 1);

            int meses_pdtes;
            int aniomes;
            DateTime fecha_registro = new DateTime();

            DateTime fecha_actual = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);

            db = new MySQLDB(MySQLDB.Esquemas.GBL);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                f++;
                c = 1;

                if (r["cd_empr"] != System.DBNull.Value)
                    workSheet.Cells[f, c].Value = r["cd_empr"].ToString();
                c++;

                if (r["de_tp_cli"] != System.DBNull.Value)
                    workSheet.Cells[f, c].Value = r["de_tp_cli"].ToString();
                c++;

                if (portugal)
                {
                    if (r["cd_tp_tension"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["cd_tp_tension"].ToString();
                    c++;
                }

                if (r["cd_nif_cif_cli"] != System.DBNull.Value)
                    workSheet.Cells[f, c].Value = r["cd_nif_cif_cli"].ToString();
                c++;

                if (r["tx_apell_cli"] != System.DBNull.Value)
                    workSheet.Cells[f, c].Value = r["tx_apell_cli"].ToString();
                c++;

                if (r["fh_alta_crto"] != System.DBNull.Value)
                {
                    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_alta_crto"]).Date;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (r["fh_inicio_vers_crto"] != System.DBNull.Value)
                {
                    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_inicio_vers_crto"]).Date;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (r["cups20"] != System.DBNull.Value)
                    workSheet.Cells[f, c].Value = r["cups20"].ToString();
                c++;

                if (r["id_instalacion"] != System.DBNull.Value)
                    workSheet.Cells[f, c].Value = r["id_instalacion"].ToString();
                c++;

                if (r["cd_tarifa_c"] != System.DBNull.Value)
                    workSheet.Cells[f, c].Value = r["cd_tarifa_c"].ToString();
                c++;

                if (r["cd_crto_comercial"] != System.DBNull.Value)
                    workSheet.Cells[f, c].Value = r["cd_crto_comercial"].ToString();
                c++;

                if (r["mes"] != System.DBNull.Value)
                {
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["mes"]);
                    aniomes = Convert.ToInt32(r["mes"]);
                    fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                        Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                }
                c++;

                if (r["de_empr_distdora_nombre"] != System.DBNull.Value)
                    workSheet.Cells[f, c].Value = r["de_empr_distdora_nombre"].ToString();
                c++;

                if (r["de_estado"] != System.DBNull.Value)
                    workSheet.Cells[f, c].Value = r["de_estado"].ToString();
                c++;

                if (r["de_subestado"] != System.DBNull.Value)
                    workSheet.Cells[f, c].Value = r["de_subestado"].ToString();
                c++;



                workSheet.Cells[f, c].Value = GetDiasEstado(r["cups20"].ToString());

                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
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
                    workSheet.Cells[f, c].Value = meses_pdtes;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                c++;


                if (r["mes"] != System.DBNull.Value && r["TAM"] != System.DBNull.Value)
                {
                    //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                    meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["tam"]) * meses_pdtes;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                }
                else
                {
                    c++;
                }

                if (r["agora"] != System.DBNull.Value)
                    workSheet.Cells[f, c].Value = r["agora"].ToString();

                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                c++;



            }
            db.CloseConnection();

            headerCells = workSheet.Cells[1, 1, 1, c];
            headerFont = headerCells.Style.Font;
            headerFont.Bold = true;
            allCells = workSheet.Cells[1, 1, f, c];

            allCells.AutoFitColumns();
        }


        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> CargaPendienteHist_DesdeFecha(DateTime f, List<string> lista_empresas)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>>();

            DateTime fecha_actual = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime fecha_registro = new DateTime();
            int meses_pdtes = 0;
            int aniomes = 0;
            int num_dias_informe = 0;
            bool firstOnly = true;

            try
            {




                //sof.Contruye_Sofisticados();
                //agoraManual = CargaAgoraManual(DateTime.Now, DateTime.Now);
                //agoraPortugal = new contratacion.Agora_Portugal();

                strSql = " SELECT fecha_informe, pend.empresa_titular AS EMPRESA,"
                    + " pend.cups13, "
                    + " pend.mes as aaaammPdte, pend.estado, pend.subestado, pend.tam"
                    + " FROM fact. pend where "
                    + " fecha_informe >= '" + f.ToString("yyyy-MM-dd") + "'"
                    + " ORDER BY pend.fecha_informe, pend.empresa_titular, "
                    + " pend.cups13, pend.mes ASC";

                strSql = "SELECT p.cd_empr_titular, ps.cd_empr, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                        + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, ps.cups20, ps.cd_tarifa_c,"
                        + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre, p.cd_estado, p.cd_subestado,"
                        + " de.de_estado, ds.de_subestado, p.fh_periodo, p.agora, p.TAM,"
                        + " p.lg_multimedida, p.fec_act"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"
                        + " where p.fec_act >= '" + DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd") + "'";

                foreach(string p in lista_empresas)
                {
                    if (firstOnly)
                    {
                        strSql += " and cd_empr_titular in ("
                            + "'" + p + "'";
                        firstOnly = false;
                    }else
                        strSql += ",'" + p + "'";
                }
                

                        
                   strSql += ") ORDER BY p.fec_act desc, ps.cd_empr, "
                        + " ps.cups20, p.fh_periodo ASC";


                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {                    
                    EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                    c.cod_empresaTitular = r["cd_empr_titular"].ToString();
                    c.empresaTitular = r["cd_empr"].ToString();                    
                    c.cups20 =  r["cups20"].ToString();
                    c.aaaammPdte = Convert.ToInt32(r["fh_periodo"]);
                    c.cod_estado = r["cd_estado"].ToString();
                    c.cod_subestado = r["cd_subestado"].ToString();
                    c.estado = r["de_estado"].ToString();
                    c.subsEstado = r["de_subestado"].ToString();
                    c.fecha_informe = Convert.ToDateTime(r["fec_act"]).Date;


                    if (r["tam"] != System.DBNull.Value)
                    {
                        if (c.aaaammPdte != 0)
                        {
                            aniomes = Convert.ToInt32(c.aaaammPdte);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);

                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12)
                                + fecha_actual.Month - fecha_registro.Month;
                            c.tam = Convert.ToDouble(r["tam"]) * meses_pdtes;
                        }
                        else
                            c.tam = Convert.ToDouble(r["tam"]);

                    }

                    if (r["agora"] != System.DBNull.Value)
                        if(r["agora"].ToString() != "N")
                            c.agora = true;
                        else
                            c.agora = false;
                    else
                        c.agora = false;

                    if (r["lg_multimedida"] != System.DBNull.Value)
                        c.multimedida = r["lg_multimedida"].ToString() == "S";
                    else
                        c.multimedida = false;


                    // Sólo necesitamos los últimos 5 días para el informe                    
                    

                    List<EndesaEntity.medida.Pendiente> o;
                    if (!d.TryGetValue(c.fecha_informe, out o))
                    {                       
                        
                        o = new List<EndesaEntity.medida.Pendiente>();
                        o.Add(c);
                        d.Add(c.fecha_informe, o);                   
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

        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> CargaPendienteHist_PT_DesdeFecha(DateTime f, List<string> lista_empresas, List<string> lista_segmentos)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>>();

            DateTime fecha_actual = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime fecha_registro = new DateTime();
            int meses_pdtes = 0;
            int aniomes = 0;
            int num_dias_informe = 0;
            bool firstOnly = true;

            try
            {

                //sof.Contruye_Sofisticados();
                //agoraManual = CargaAgoraManual(DateTime.Now, DateTime.Now);
                //agoraPortugal = new contratacion.Agora_Portugal();
                

                strSql = "SELECT p.cd_empr_titular, ps.cd_empr, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                        + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, ps.cups20, ps.cd_tarifa_c, ps.cd_tp_tension,"
                        + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre, p.cd_estado, p.cd_subestado,"
                        + " de.de_estado, ds.de_subestado, p.fh_periodo, p.agora, p.TAM,"
                        + " p.lg_multimedida, p.fec_act"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps_pt ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"
                        + " where p.fec_act >= '" + DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd") + "'";

                foreach (string p in lista_empresas)
                {
                    if (firstOnly)
                    {
                        strSql += " and cd_empr_titular in ("
                            + "'" + p + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + p + "'";
                }

                firstOnly = true;
                foreach (string p in lista_segmentos)
                {
                    if (firstOnly)
                    {
                        strSql += ") and cd_tp_tension in ("
                            + "'" + p + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + p + "'";
                }



                strSql += ") ORDER BY p.fec_act desc, ps.cd_empr, "
                     + " ps.cups20, p.fh_periodo ASC";


                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                    c.cod_empresaTitular = r["cd_empr_titular"].ToString();
                    c.empresaTitular = r["cd_empr"].ToString();
                    c.cups20 = r["cups20"].ToString();
                    c.aaaammPdte = Convert.ToInt32(r["fh_periodo"]);
                    c.cod_estado = r["cd_estado"].ToString();
                    c.cod_subestado = r["cd_subestado"].ToString();
                    c.estado = r["de_estado"].ToString();
                    c.subsEstado = r["de_subestado"].ToString();
                    c.fecha_informe = Convert.ToDateTime(r["fec_act"]).Date;

                    if (r["cd_tp_tension"] != System.DBNull.Value)
                        c.segmento = r["cd_tp_tension"].ToString();


                    if (r["tam"] != System.DBNull.Value)
                    {
                        if (c.aaaammPdte != 0)
                        {
                            aniomes = Convert.ToInt32(c.aaaammPdte);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);

                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12)
                                + fecha_actual.Month - fecha_registro.Month;
                            c.tam = Convert.ToDouble(r["tam"]) * meses_pdtes;
                        }
                        else
                            c.tam = Convert.ToDouble(r["tam"]);

                    }

                    if (r["agora"] != System.DBNull.Value)
                        if (r["agora"].ToString() != "N")
                            c.agora = true;
                        else
                            c.agora = false;
                    else
                        c.agora = false;

                    //if (r["lg_multimedida"] != System.DBNull.Value)
                    //    c.multimedida = r["lg_multimedida"].ToString() == "S";
                    //else
                    //    c.multimedida = false;



                    // Sólo necesitamos los últimos 5 días para el informe                    


                    List<EndesaEntity.medida.Pendiente> o;
                    if (!d.TryGetValue(c.fecha_informe, out o))
                    {

                        o = new List<EndesaEntity.medida.Pendiente>();
                        o.Add(c);
                        d.Add(c.fecha_informe, o);
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

        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaTotales(List<string> lista_empresas)
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
                strSql = "SELECT t.fh_envio, t.cd_estado, t.cd_subestado, t.num_cups, t.tam"
                    + " FROM t_ed_h_sap_pendiente_facturar_agrupado_totales t"
                    + " where cd_empr_titular in ('" + lista_empresas[0] + "'";

                for (int x = 1; x < lista_empresas.Count; x++)
                    strSql += ",'" + lista_empresas[x] + "'";

                strSql += ") ORDER BY t.fh_envio DESC, t.cd_estado";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha_informe = Convert.ToDateTime(r["fh_envio"]);
                    estado = r["cd_estado"].ToString();
                    subestado = r["cd_subestado"].ToString();
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

        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaAgora_TAM(DateTime fd, List<string> lista_empresas)
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
                strSql = "SELECT t.fh_envio, t.cd_estado, t.cd_subestado, t.num_cups, t.tam"
                     + " FROM t_ed_h_sap_pendiente_facturar_agrupado_totales t"
                     + " where cd_empr_titular in ('" + lista_empresas[0] + "'";

                for (int x = 1; x < lista_empresas.Count; x++)
                    strSql += ",'" + lista_empresas[x] + "'";

                strSql += ") and t.agora = 'S'" 
                    + " and t.fh_envio >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " ORDER BY t.fh_envio, t.cd_estado";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    fecha_informe = Convert.ToDateTime(r["fh_envio"]);
                    estado = r["cd_estado"].ToString();
                    subestado = r["cd_subestado"].ToString();
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

        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaNoAgora_TAM(DateTime fd, List<string> lista_empresas)
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
                strSql = "SELECT t.fh_envio, t.cd_estado, t.cd_subestado, t.num_cups, t.tam"
                      + " FROM t_ed_h_sap_pendiente_facturar_agrupado_totales t"
                      + " where cd_empr_titular in ('" + lista_empresas[0] + "'";

                for (int x = 1; x < lista_empresas.Count; x++)
                    strSql += ",'" + lista_empresas[x] + "'";

                strSql += ") and t.agora = 'N'"
                    + " and t.fh_envio >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " ORDER BY t.fh_envio, t.cd_estado";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha_informe = Convert.ToDateTime(r["fh_envio"]);
                    estado = r["cd_estado"].ToString();
                    subestado = r["cd_subestado"].ToString();
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
            

            Pendiente_Estados estados = new Pendiente_Estados();
            Pendiente_Subestados subestados = new Pendiente_Subestados();


            foreach(KeyValuePair<string, EndesaEntity.medida.Pendiente> p in subestados.dic)
            {
                string[] texto = p.Key.Split('.');

                EndesaEntity.medida.Pendiente_Totales c = new EndesaEntity.medida.Pendiente_Totales();
                c.estado = texto[0];
                c.subestado = p.Key;
                t.Add(c);
            }
                       

            return t;
        }

        private int Total_Pendiente(bool agora, DateTime fecha, List<string> lista_empresas, List<string> lista_segmentos, string estado, string subestado)
        {
            int total = 0;

            List<EndesaEntity.medida.Pendiente> o;

            if (lista_segmentos == null)
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int i = 0; i < o.Count; i++)
                        {
                            if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora &&
                                (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                total = total + 1;

                        }
                }
            }
            else
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int z = 0; z < lista_segmentos.Count; z++)
                            for (int i = 0; i < o.Count; i++)
                            {
                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z]
                                    && o[i].agora == agora &&
                                    (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                    total = total + 1;

                            }
                }
            }                

            return total;

        }
              


        private double Total_Pendiente_TAM(bool agora, DateTime fecha, List<string> lista_empresas, List<string> lista_segmentos, string estado, string subestado)
        {
            double total = 0;

            List<EndesaEntity.medida.Pendiente> o;

            if(lista_segmentos == null)
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int i = 0; i < o.Count; i++)
                        {
                            if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora &&
                                (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                total = total + o[i].tam;

                        }


                }
            }                
            else
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int z = 0; z < lista_segmentos.Count; z++)
                            for (int i = 0; i < o.Count; i++)
                        {
                            if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z]
                                    && o[i].agora == agora &&
                                (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                total = total + o[i].tam;

                        }


                }
            }

            return total;

        }

        private string DetalleExcel(DateTime fecha_informe, List<string> lista_empresas, List<string> lista_tension)
        {
            string strSql = "";
            bool firstOnly = true;
            
            strSql = "SELECT ps.cd_empr, ps.cd_tp_tension, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, p.cd_cups as cups20, p.id_instalacion, ps.cd_tarifa_c,"
                + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre, p.lg_multimedida,"
                + " concat(p.cd_estado,' ',de.de_estado) as de_estado, concat(p.cd_subestado,' ',if (ds.de_subestado is null,'', ds.de_subestado)) as de_subestado, p.fh_periodo as mes, p.agora, p.TAM"
                + " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p";

            if (lista_empresas[0].Contains("PT"))
                strSql += " LEFT OUTER JOIN cont.t_ed_h_ps_pt ps ON";
            else
                strSql += " LEFT OUTER JOIN cont.t_ed_h_ps ps ON";

            strSql += " ps.cups20 = p.cd_cups"
                + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                + " de.cd_estado = p.cd_estado"
                + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                + " ds.cd_subestado = p.cd_subestado"
                + " where p.fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "' and"
                + " p.cd_empr_titular in (";
            
            foreach(string p in lista_empresas)
            {
                if (firstOnly)
                {
                    strSql += "'" + p + "'";
                    firstOnly = false;
                }
                else
                    strSql += ",'" + p + "'";

            }

            strSql += ")";

            if(lista_tension != null)
            {
                firstOnly = true;
                strSql += " and ps.cd_tp_tension in (";
                foreach (string p in lista_tension)
                {
                    if (firstOnly)
                    {
                        strSql += "'" + p + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + p + "'";

                }
                strSql += ")";
            }

            return strSql;

        }

        private void EnvioCorreo_PdteWeb_BI(string archivo, DateTime fecha_informe)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                string from = param.GetValue("mail_from_psat_tam");
                string to = param.GetValue("mail_to_psat_tam");
                string cc = param.GetValue("mail_cc_psat_tam");
                string subject = param.GetValue("mail_subject_psat_tam") + " a " + fecha_informe.ToString("dd/MM/yyyy");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Se adjunta el archivo ").Append(fileInfo.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");

                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (param.GetValue("mail_enviar_mail_psat_tam") == "S")
                    mes.SendMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);

                else
                    mes.SaveMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("mail_from"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }

        private DateTime CalculaFechaDesdeinforme()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime date = DateTime.Now;

            strSql = "SELECT c.fh_envio"
                + " FROM t_ed_h_sap_pendiente_facturar_agrupado c"
                + " GROUP BY c.fh_envio desc"
                + " LIMIT 5";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["fh_envio"] != System.DBNull.Value)
                    date = Convert.ToDateTime(r["fh_envio"]);
            }
            db.CloseConnection();

            return date;

        }

        private DateTime CalculaFechaDetalle()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime date = DateTime.Now;

            strSql = "SELECT max(c.fh_envio) as max_fecha"
                + " FROM t_ed_h_sap_pendiente_facturar_agrupado c";
                

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["max_fecha"] != System.DBNull.Value)
                    date = Convert.ToDateTime(r["max_fecha"]);
            }
            db.CloseConnection();

            return date;

        }

        private Dictionary<string, DateTime> CargaDiasEstado()
        {
            Dictionary<string, DateTime> d = new Dictionary<string, DateTime>();

           MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string cups = "";
            DateTime primera_fecha = DateTime.Now;

            strSql = "SELECT p.cd_cups, p.fh_periodo, MIN(p.fh_envio) AS primera_fecha, p.cd_subestado"
                + " FROM t_ed_h_sap_pendiente_facturar_agrupado p"
                + " WHERE p.fh_periodo >= " + DateTime.Now.AddYears(-1).ToString("yyyyMM") 
                + " GROUP BY p.cd_cups, p.fh_periodo, p.cd_subestado"
                + " ORDER BY p.fh_envio DESC";


            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                cups = r["cd_cups"].ToString();
                primera_fecha = Convert.ToDateTime(r["primera_fecha"]);

                DateTime o;
                if (!d.TryGetValue(cups, out o))
                    d.Add(cups, primera_fecha);

            }
            db.CloseConnection();
            return d;

        }

        private int GetDiasEstado(string cups)
        {
            int dias = 1;
            DateTime o;
            if (dic_dias_estado.TryGetValue(cups, out o))
            {
                dias = (DateTime.Now.Date - o.Date).Days + 1;
            }
                
            

            return dias;
        }

        private Dictionary<string, EndesaEntity.medida.Pendiente> Carga()
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";


            Dictionary<string, EndesaEntity.medida.Pendiente> d =
                new Dictionary<string, EndesaEntity.medida.Pendiente>();


            try
            {
                strSql = "SELECT p.cd_cups,"
                    + " concat(p.cd_estado, ' ', e.de_estado) as de_estado,"
                    + " concat(p.cd_subestado, ' ', s.de_subestado) as de_subestado,"
                    + " p.fh_periodo as mes"
                    + " FROM t_ed_h_sap_pendiente_facturar_agrupado p"
                    + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar e ON"
                    + " e.cd_estado = p.cd_estado"
                    + " LEFT OUTER JOIN t_ed_p_subestado_sap_pendiente_facturar s ON"
                    + " s.cd_subestado = p.cd_subestado"
                    + " WHERE p.fh_envio = (SELECT MAX(fh_envio) AS max_fh_envio FROM t_ed_h_sap_pendiente_facturar_agrupado)";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                    if (r["cd_cups"] != System.DBNull.Value)
                        c.cups20 = r["cd_cups"].ToString();

                    if (r["de_estado"] != System.DBNull.Value)
                        c.descripcion_estado = r["de_estado"].ToString();

                    if (r["de_subestado"] != System.DBNull.Value)
                        c.descripcion_subestado = r["de_subestado"].ToString();

                    if (r["mes"] != System.DBNull.Value)
                        c.aaaammPdte = Convert.ToInt32(r["mes"]);

                    EndesaEntity.medida.Pendiente o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception ex)
            {
                return null;
            }


        }

        public void GetEstados(string cups20)
        {
            EndesaEntity.medida.Pendiente o;
            if (dic_pendiente.TryGetValue(cups20, out o))
            {
                this.existe = true;
                this.cups20 = o.cups20;
                this.descripcion_estado = o.descripcion_estado;
                this.descripcion_subestado = o.descripcion_subestado;
                this.aaaammPdte = o.aaaammPdte;
            }
            else
                this.existe = false;
        }
    }
}

