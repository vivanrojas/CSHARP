using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.utilidades;
using EndesaBusiness.servidores;
using OfficeOpenXml;
using EndesaBusiness.contratacion;
using OfficeOpenXml.Style;
using System.Globalization;
using System.IO;

namespace EndesaBusiness.medida
{
    public class Pend_Informe_FAC
    {

        // Se construye este informe a peticion de 
        // Ignacio Villar.

        utilidades.Param param;
        logs.Log ficheroLog;
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;

        contratacion.ComplementosContrato complementosContratoPS;
        contratacion.Agora_Portugal agoraPortugal;

        Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> dic_pendiente_hist_fecha;


        EndesaBusiness.facturacion.cuadro_mando.Sofisticados sof =
                new facturacion.cuadro_mando.Sofisticados();

        Dictionary<string, EndesaEntity.facturacion.AgoraManual> agoraManual;
        public Pend_Informe_FAC()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Informe_Pend");
            param = new utilidades.Param("dt_vw_ed_f_detalle_pendiente_facturar_param", servidores.MySQLDB.Esquemas.MED);
            ss_pp = new utilidades.Seguimiento_Procesos();
        }

        public void GeneraInformePendWeb_PSAT_TAM(bool automatico)
        {
            FileInfo file;
            string ruta_salida_archivo = "";

            string[] listaArchivos = Directory.GetFiles(param.GetValue("ruta_salida_informe"),
                    param.GetValue("prefijo_informe") + "*.xlsx");
            for (int i = 0; i < listaArchivos.Length; i++)
            {
                file = new FileInfo(listaArchivos[i]);
                file.Delete();
            }

            ruta_salida_archivo = param.GetValue("ruta_salida_informe")
            + param.GetValue("prefijo_informe")
            + DateTime.Now.ToString("yyyyMMdd_HHmmss")
            + param.GetValue("sufijo_informe");

            InformePendWeb_PSAT_TAM_V2(ruta_salida_archivo, automatico);
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
                    if (!d.TryGetValue(c.cups13, out o))
                        d.Add(c.cups13, c);

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        private bool EsAgoraManual(Dictionary<string, EndesaEntity.facturacion.AgoraManual> dic, string cups13)
        {
            
            EndesaEntity.facturacion.AgoraManual o;
            return dic.TryGetValue(cups13, out o);
        }
        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> CargaPendienteHist_DesdeFecha(DateTime f)
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
            


            try
            {

                sof.Contruye_Sofisticados();
                agoraManual = CargaAgoraManual(DateTime.Now, DateTime.Now);
                agoraPortugal = new contratacion.Agora_Portugal();

                strSql = " SELECT fecha_informe, pend.empresa_titular AS EMPRESA,"
                    + " pend.cups13, "
                    + " pend.mes as aaaammPdte, pend.estado, pend.subestado, pend.tam"
                    + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar_agrupado_hist_t pend where "
                    + " fecha_informe >= '" + f.ToString("yyyy-MM-dd") + "'"
                    // Este CUPS se quita a petición de Ignacio Villar
                    + " and pend.cups13 <> 'VZZ0605131910'"
                    + " ORDER BY pend.fecha_informe, pend.empresa_titular, "
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
                    c.fecha_informe = Convert.ToDateTime(r["fecha_informe"]);

                    if (r["tam"] != System.DBNull.Value)
                    {
                        if(c.aaaammPdte != 0)
                        {
                            aniomes = Convert.ToInt32(c.aaaammPdte);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);

                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) 
                                + fecha_actual.Month - fecha_registro.Month;
                            c.tam = Convert.ToDouble(r["tam"]) * meses_pdtes;
                        }else
                            c.tam = Convert.ToDouble(r["tam"]);

                    }
                        

                    c.agora = sof.EsSofisticado(c.cups13) || EsAgoraManual(agoraManual, c.cups13);


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

        private int Total_Pendiente(bool agora, DateTime fecha, List<string> lista_empresas, string estado, string subestado)
        {
            int total = 0;

            List<EndesaEntity.medida.Pendiente> o;
            if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
            {
                for (int j = 0; j < lista_empresas.Count; j++)
                    for (int i = 0; i < o.Count; i++)
                    {
                        if (o[i].empresaTitular == lista_empresas[j] && o[i].agora == agora && 
                            (o[i].estado == estado && (o[i].subsEstado == subestado || subestado == null)))
                            total = total + 1;               
                        
                    }
                    
                
            }

            return total;

        }
        private double Total_Pendiente_TAM(bool agora, DateTime fecha, List<string> lista_empresas, string estado, string subestado)
        {
            double total = 0;

            List<EndesaEntity.medida.Pendiente> o;
            if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
            {
                for (int j = 0; j < lista_empresas.Count; j++)
                    for (int i = 0; i < o.Count; i++)
                    {
                        if (o[i].empresaTitular == lista_empresas[j] && o[i].agora == agora &&
                            (o[i].estado == estado && (o[i].subsEstado == subestado || subestado == null)))
                        {
                            total = total + o[i].tam;
                        }
                            

                    }


            }

            return total;

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

            bool tiene_complemento_a01 = false;
            bool sacar_portugal = true;

            //Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_totales;
            //Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_agora;
            //Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_Noagora;
            //Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_agora_tam;
            //Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> dic_Noagora_tam;



            utilidades.Fechas utilfecha = new Fechas();

            DateTime fd = new DateTime();
            DateTime fd_tam = new DateTime();
            DateTime udh = new DateTime();

            try
            {
                // Ruta de la plantilla

                // Este proceso tiene como dependencia el proceso de medida 
                // ZZ_MED_COPIA_PENDIENTE_WEB_REDSHIFT por eso preguntamos por el
                if (ss_pp.GetFecha_FinProceso("Medida", "PENDIENTE", "COPIA PENDIENTE REDSHIFT").Date == DateTime.Now.Date)
                {

                    if (automatico && (ss_pp.GetFecha_FinProceso("Facturación", "PendML", "Agrupado Totales").Date < DateTime.Now.Date))
                    {


                        ss_pp.Update_Fecha_Inicio("Facturación", "PendML", "Agrupado Totales");

                        FileInfo plantillaExcel = new FileInfo(System.Environment.CurrentDirectory + param.GetValue("plantilla_PS_AT_TAM"));

                        FileInfo fileInfo = new FileInfo(ruta_salida_archivo);
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                        ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);

                        // #region Resumen

                        var workSheet = excelPackage.Workbook.Worksheets["Resumen ES"];
                        var headerCells = workSheet.Cells[1, 1, 1, 17];
                        var headerFont = headerCells.Style.Font;


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


                        dic_pendiente_hist_fecha = CargaPendienteHist_DesdeFecha(fd);

                        fd_tam = utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil());

                        List<string> lista_empresas_ES = new List<string>();
                        lista_empresas_ES.Add("20");
                        lista_empresas_ES.Add("70");


                        int totales_dia = 0;
                        double totales_dia_tam = 0;

                        bool noAgora = false;
                        bool siAgora = true;

                        int totales_dia_medida_no_agora = 0;
                        double totales_dia_medida_no_agora_tam = 0;

                        int totales_dia_medida_agora = 0;
                        double totales_dia_medida_agora_tam = 0;

                        int totales_dia_facturacion_no_agora = 0;
                        double totales_dia_facturacion_no_agora_tam = 0;

                        int totales_dia_facturacion_agora = 0;
                        double totales_dia_facturacion_agora_tam = 0;


                        Dictionary<DateTime, int> dic_Totales_cups = new Dictionary<DateTime, int>();
                        Dictionary<DateTime, double> dic_Totales_tam = new Dictionary<DateTime, double>();



                        int dia = 0;
                        int dia_tam = 0;


                        c = 3;




                        // NO ÁGORA
                        foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                        {


                            dia++;

                            Console.WriteLine("Totales ES noAgora dia: " + p.Key.ToString("dd/MM/yyyy"));


                            f = 4;
                            c++;

                            workSheet.Cells[f, c].Value = p.Key;
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.B. Endesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "05. CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.A. No Validada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.B. Validada con Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.C. Validada sin Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.E. Modificada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                            totales_dia_medida_no_agora =
                                  Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "02. CC Rechazada por CS", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "03. CC Completa en CS", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "04. CC Enviada a SCE ML", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "05. CC Rechazada por SCE ML", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "05. CC Enviada a SCE ML", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "06. CC Incompleta  SCE ML", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "07. LTP SCE", null);


                            workSheet.Cells[f, c].Value = totales_dia_medida_no_agora;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", "08.B. Bloqueado");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", "08.C. Otros");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                            totales_dia_facturacion_no_agora = Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "09. El Punto está Extraído", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_ES, "10. Prefactura pendiente", null);

                            workSheet.Cells[f, c].Value = totales_dia_facturacion_no_agora;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;


                            // total No Ágora
                            workSheet.Cells[f, c].Value = totales_dia_medida_no_agora + totales_dia_facturacion_no_agora;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, c].Style.Font.Bold = true;

                            int o;
                            if (!dic_Totales_cups.TryGetValue(p.Key, out o))
                                dic_Totales_cups.Add(p.Key, totales_dia_medida_no_agora + totales_dia_facturacion_no_agora);

                        }

                        c = 8;


                        // NO ÁGORA TAM
                        foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                        {


                            dia_tam++;

                            #region diferencia

                            Console.WriteLine("Totales TAM ES noAgora dia: " + p.Key.ToString("dd/MM/yyyy"));

                            #endregion

                            f = 4;
                            c++;

                            workSheet.Cells[f, c].Value = p.Key;
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.B. Endesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.A. No Validada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.B. Validada con Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.C. Validada sin Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.E. Modificada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;


                            totales_dia_medida_no_agora_tam = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "02. CC Rechazada por CS", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "03. CC Completa en CS", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "04. CC Enviada a SCE ML", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "05.CC Rechazada por SCE ML", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "05. CC Enviada a SCE ML", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "06. CC Incompleta  SCE ML", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "07. LTP SCE", null);


                            workSheet.Cells[f, c].Value = totales_dia_medida_no_agora_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", "08.B. Bloqueado");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", "08.C. Otros");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            totales_dia_facturacion_no_agora_tam = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "09. El Punto está Extraído", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_ES, "10. Prefactura pendiente", null);

                            workSheet.Cells[f, c].Value = totales_dia_facturacion_no_agora_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = totales_dia_medida_no_agora_tam + totales_dia_facturacion_no_agora_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            workSheet.Cells[f, c].Style.Font.Bold = true;


                            double o;
                            if (!dic_Totales_tam.TryGetValue(p.Key, out o))
                                dic_Totales_tam.Add(p.Key, totales_dia_medida_no_agora_tam + totales_dia_facturacion_no_agora_tam);

                        }


                        headerCells = workSheet.Cells[1, 1, 1, 30];
                        headerFont = headerCells.Style.Font;
                        headerFont.Bold = true;
                        var allCells = workSheet.Cells[1, 1, 50, 50];


                        c = 3;
                        dia = 0;


                        // ÁGORA
                        foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                        {

                            f = 26;
                            c++;


                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.B. Endesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.A. No Validada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.B. Validada con Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.C. Validada sin Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.E. Modificada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                            totales_dia_medida_agora = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "02. CC Rechazada por CS", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "03. CC Completa en CS", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "04. CC Enviada a SCE ML", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "05.CC Rechazada por SCE ML", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "05. CC Enviada a SCE ML", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "06. CC Incompleta  SCE ML", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "07. LTP SCE", null);


                            workSheet.Cells[f, c].Value = totales_dia_medida_agora;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", "08.B. Bloqueado");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", "08.C. Otros");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                            totales_dia_facturacion_agora = Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "09. El Punto está Extraído", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_ES, "10. Prefactura pendiente", null);

                            workSheet.Cells[f, c].Value = totales_dia_facturacion_agora;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = totales_dia_medida_agora + totales_dia_facturacion_agora;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            int o;
                            if (dic_Totales_cups.TryGetValue(p.Key, out o))
                            {
                                workSheet.Cells[f, c].Value = o +
                               totales_dia_medida_agora + totales_dia_facturacion_agora;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                            }


                            f++;


                        }

                        c = 8;

                        dia_tam = 0;


                        // ÁGORA TAM
                        foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                        {



                            f = 25;
                            c++;

                            f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.B. Endesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.A. No Validada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.B. Validada con Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.C. Validada sin Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "07. LTP SCE", "07.E. Modificada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;


                            totales_dia_medida_agora_tam = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "01. Pendiente de medida", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "02. CC Rechazada por CS", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "03. CC Completa en CS", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "04. CC Enviada a SCE ML", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "05. CC Rechazada por SCE ML", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "05. CC Enviada a SCE ML", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "06. CC Incompleta  SCE ML", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "07. LTP SCE", null);


                            workSheet.Cells[f, c].Value = totales_dia_medida_agora_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", "08.B. Bloqueado");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", "08.C. Otros");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            totales_dia_facturacion_agora_tam = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "08. El Punto no está Extraído", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "09. El Punto está Extraído", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_ES, "10. Prefactura pendiente", null);

                            workSheet.Cells[f, c].Value = totales_dia_facturacion_agora_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = totales_dia_medida_agora_tam + totales_dia_facturacion_agora_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            workSheet.Cells[f, c].Style.Font.Bold = true;

                            double o;
                            if (dic_Totales_tam.TryGetValue(p.Key, out o))
                            {
                                workSheet.Cells[47, c].Value = o +
                                totales_dia_medida_agora_tam +
                                totales_dia_facturacion_agora_tam;
                                workSheet.Cells[47, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[47, c].Style.Font.Bold = true;
                            }




                        }

                        c = 3;
                        //for (int j = 0; j < 5; j++)
                        //{
                        //    c++;
                        //    workSheet.Cells[47, c].Value = total_general[j];
                        //    workSheet.Cells[47, c].Style.Numberformat.Format = "#,##0";
                        //    workSheet.Cells[47, c].Style.Font.Bold = true;

                        //}

                        c = 8;
                        //for (int j = 0; j < 2; j++)
                        //{
                        //    c++;
                        //    workSheet.Cells[47, c].Value = total_general_tam[j];
                        //    workSheet.Cells[47, c].Style.Numberformat.Format = "#,##0.00";
                        //    workSheet.Cells[47, c].Style.Font.Bold = true;

                        //}






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
                            + " contrato, mes, distribuidora, estado, subestado, multimedida, if(p.tam is null, 0, p.tam) as tam "
                            + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar_agrupado_t p"
                            + " LEFT OUTER JOIN cont.PS_AT ps ON"
                            + " ps.IDU = p.cups13"
                            + " WHERE p.fecha_informe = '" + udh.ToString("yyyy-MM-dd") + "' and"
                            + " p.empresa_titular <> 80"
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

                            if (r["tam"] != System.DBNull.Value)
                            {
                                workSheet.Cells[f, c].Value = Convert.ToDouble(r["tam"]);
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            }

                            c++;

                            if (r["mes"] != System.DBNull.Value)
                            {
                                //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                                meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                                workSheet.Cells[f, c].Value = meses_pdtes; c++;
                                workSheet.Cells[f, c].Value = Convert.ToDouble(r["tam"]) * meses_pdtes;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                            }
                            else
                            {
                                c++; c++;
                            }

                            // AGORA                   

                            if (r["cups13"] != System.DBNull.Value)
                            {
                                tiene_complemento_a01 = sof.EsSofisticado(r["cups13"].ToString()) || EsAgoraManual(agoraManual, r["cups13"].ToString());


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

                        #endregion

                        dic_Totales_cups.Clear();
                        dic_Totales_tam.Clear();

                        #region Resumen Portugal MT
                        List<string> lista_empresas_PT = new List<string>();
                        lista_empresas_PT.Add("80");

                        //dic_totales = CargaTotales(lista_empresas_PT);
                        //dic_agora = CargaAgora_TAM(fd, lista_empresas_PT);
                        //dic_Noagora = CargaNoAgora_TAM(fd, lista_empresas_PT);

                        //dic_agora_tam = CargaAgora_TAM(fd_tam, lista_empresas_PT);
                        //dic_Noagora_tam = CargaNoAgora_TAM(fd_tam, lista_empresas_PT);

                        workSheet = excelPackage.Workbook.Worksheets["Resumen Portugal MT"];


                        totales_dia = 0;
                        totales_dia_tam = 0;




                        dia = 0;
                        dia_tam = 0;


                        c = 3;


                        // No AGORA MT
                        foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                        {

                            f = 4;
                            c++;

                            workSheet.Cells[f, c].Value = p.Key;
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.B. Endesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.A. No Validada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.B. Validada con Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.C. Validada sin Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.E.Modificada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                            totales_dia_medida_no_agora = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "02. CC Rechazada por CS", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "03. CC Completa en CS", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "04. CC Enviada a SCE ML", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "05.CC Rechazada por SCE ML", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "05. CC Enviada a SCE ML", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "06. CC Incompleta  SCE ML", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "07. LTP SCE", null);


                            workSheet.Cells[f, c].Value = totales_dia_medida_no_agora;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", "08.B. Bloqueado");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", "08.C. Otros");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                            totales_dia_facturacion_no_agora = Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "09. El Punto está Extraído", null)
                                + Total_Pendiente(noAgora, p.Key, lista_empresas_PT, "10. Prefactura pendiente", null);

                            workSheet.Cells[f, c].Value = totales_dia_facturacion_no_agora;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = totales_dia_medida_no_agora + totales_dia_facturacion_no_agora;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, c].Style.Font.Bold = true;


                            int o;
                            if (!dic_Totales_cups.TryGetValue(p.Key, out o))
                                dic_Totales_cups.Add(p.Key, totales_dia_medida_no_agora + totales_dia_facturacion_no_agora);

                        }

                        c = 8;

                        // No Agora MT TAM
                        foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                        {


                            f = 4;
                            c++;

                            workSheet.Cells[f, c].Value = p.Key;
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.B. Endesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.A. No Validada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.B. Validada con Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.C. Validada sin Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.E. Modificada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;


                            totales_dia_medida_no_agora_tam = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "02. CC Rechazada por CS", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "03. CC Completa en CS", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "04. CC Enviada a SCE ML", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "05.CC Rechazada por SCE ML", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "05. CC Enviada a SCE ML", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "06. CC Incompleta  SCE ML", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "07. LTP SCE", null);


                            workSheet.Cells[f, c].Value = totales_dia_medida_no_agora_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", "08.B. Bloqueado");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", "08.C. Otros");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            totales_dia_facturacion_no_agora_tam = Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "09. El Punto está Extraído", null)
                                + Total_Pendiente_TAM(noAgora, p.Key, lista_empresas_PT, "10. Prefactura pendiente", null);

                            workSheet.Cells[f, c].Value = totales_dia_facturacion_no_agora_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = totales_dia_medida_no_agora_tam + totales_dia_facturacion_no_agora_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            workSheet.Cells[f, c].Style.Font.Bold = true;

                            double o;
                            if (!dic_Totales_tam.TryGetValue(p.Key, out o))
                                dic_Totales_tam.Add(p.Key, totales_dia_medida_no_agora_tam + totales_dia_facturacion_no_agora_tam);

                        }


                        headerCells = workSheet.Cells[1, 1, 1, 30];
                        headerFont = headerCells.Style.Font;
                        headerFont.Bold = true;
                        allCells = workSheet.Cells[1, 1, 50, 50];


                        c = 3;


                        // Agora MT
                        foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                        {

                            f = 26;
                            c++;


                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.B. Endesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.A. No Validada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.B. Validada con Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.C. Validada sin Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.E. Modificada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                            totales_dia_medida_agora = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "02. CC Rechazada por CS", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "03. CC Completa en CS", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "04. CC Enviada a SCE ML", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "05.CC Rechazada por SCE ML", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "05. CC Enviada a SCE ML", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "06. CC Incompleta  SCE ML", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "07. LTP SCE", null);


                            workSheet.Cells[f, c].Value = totales_dia_medida_agora;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", "08.B. Bloqueado");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", "08.C. Otros");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; f++;


                            totales_dia_facturacion_agora = Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "09. El Punto está Extraído", null)
                                + Total_Pendiente(siAgora, p.Key, lista_empresas_PT, "10. Prefactura pendiente", null);

                            workSheet.Cells[f, c].Value = totales_dia_facturacion_agora;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = totales_dia_medida_agora + totales_dia_facturacion_agora;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            int o;
                            if (dic_Totales_cups.TryGetValue(p.Key, out o))
                            {
                                workSheet.Cells[f, c].Value = o +
                               totales_dia_medida_agora + totales_dia_facturacion_agora;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                            }



                        }

                        c = 8;

                        // Agora MT TAM
                        foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                        {


                            f = 25;
                            c++;

                            f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.A. Endesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.B. Endesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.D. NoEndesa - Telemedida");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", "01.E. NoEndesa - TPL");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "02. CC Rechazada por CS", "02.A. El Punto no existe en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "03. CC Completa en CS", "03.A. CC Completa en el CS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "04. CC Enviada a SCE ML", "04.A. CC Enviada a SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "05.CC Rechazada por SCE ML", "05.A. CC Rechazada por SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "06. CC Incompleta  SCE ML", "06.A. CC Incompleta  SCE ML");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.A. No Validada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.B. Validada con Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.C. Validada sin Anomalías");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "07. LTP SCE", "07.E. Modificada");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;


                            totales_dia_medida_agora_tam = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "01. Pendiente de medida", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "02. CC Rechazada por CS", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "03. CC Completa en CS", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "04. CC Enviada a SCE ML", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "05. CC Rechazada por SCE ML", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "05. CC Enviada a SCE ML", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "06. CC Incompleta  SCE ML", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "07. LTP SCE", null);


                            workSheet.Cells[f, c].Value = totales_dia_medida_agora_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", "08.A. Facturado en el día anterior");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", "08.B. Bloqueado");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;
                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", "08.C. Otros");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "09. El Punto está Extraído", "09.A. El Punto está Extraído");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "10. Prefactura pendiente", "10.A. Prefactura pendiente");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; f++;

                            totales_dia_facturacion_agora_tam = Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "08. El Punto no está Extraído", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "09. El Punto está Extraído", null)
                                + Total_Pendiente_TAM(siAgora, p.Key, lista_empresas_PT, "10. Prefactura pendiente", null);

                            workSheet.Cells[f, c].Value = totales_dia_facturacion_agora_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            workSheet.Cells[f, c].Value = totales_dia_medida_agora_tam + totales_dia_facturacion_agora_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            f++;

                            double o;
                            if (dic_Totales_tam.TryGetValue(p.Key, out o))
                            {
                                workSheet.Cells[47, c].Value = o +
                                totales_dia_medida_agora_tam +
                                totales_dia_facturacion_agora_tam;
                                workSheet.Cells[47, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[47, c].Style.Font.Bold = true;
                            }




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
                                + " contrato, mes, distribuidora, estado, subestado, multimedida, if (p.tam IS NULL, 0, p.tam) AS TAM"
                                + " FROM med.dt_vw_ed_f_detalle_pendiente_facturar_agrupado_t p"
                                + " LEFT OUTER JOIN cont.contratos_ps_mt ps ON"
                                + " ps.CUPS = p.cups13"
                                + " WHERE p.fecha_informe = '" + udh.ToString("yyyy-MM-dd") + "' AND"
                                + " p.empresa_titular = 80 and "
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
                                            //if (agoraManual.TryGetValue(r["CUPS20"].ToString(), out o))
                                            //    tiene_complemento_a01 = true;
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

                        if (automatico && param.GetValue("mail_enviar_mail_psat_tam") == "S")
                        {
                            ss_pp.Update_Fecha_Fin("Facturación", "PendML", "Agrupado Totales");
                            EnvioCorreo_PdteWeb_PS_AT_TAM_v2(ruta_salida_archivo);
                        }

                    }
                }
                else
                {
                    ss_pp.Update_Comentario("Facturación", "PendML", "Agrupado Totales",
                        "Todavía no está disponible el Pdte Web >>> ZZ_MED_COPIA_PENDIENTE_WEB_REDSHIFT");
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ss_pp.Update_Comentario("Facturación", "PendML", "Agrupado Totales", e.Message);

            }
        }

        private void EnvioCorreo_PdteWeb_PS_AT_TAM_v2(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                string from = param.GetValue("mail_from_psat_tam");
                string to = param.GetValue("mail_to_psat_tam");
                string cc = param.GetValue("mail_cc_psat_tam");
                string subject = param.GetValue("mail_subject_psat_tam") + " a " + DateTime.Now.ToString("dd/MM/yyyy");

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

    }
}
