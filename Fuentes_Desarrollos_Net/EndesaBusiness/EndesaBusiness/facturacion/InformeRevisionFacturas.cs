using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using EndesaEntity.global;
using EndesaEntity.medida;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MySQLDB = EndesaBusiness.servidores.MySQLDB;

namespace EndesaBusiness.facturacion
{
    public class InformeRevisionFacturas
    {
        logs.Log ficheroLog;
        utilidades.Param param;
        List<EndesaEntity.facturacion.InformeRevisionFacturasTablaDias> lista_tabla_dias_Agora;
        List<EndesaEntity.facturacion.InformeRevisionFacturasTablaDias> lista_tabla_dias_NoAgora;
        EndesaBusiness.contratacion.ComplementosContrato complementosATR_AGORA;
        EndesaBusiness.contratacion.ComplementosContrato complementosATR_PPH; // Pass Pool Horario
        EndesaBusiness.contratacion.ComplementosContrato complementosATR_PPP; // Pass Pool x Periodo 
        EndesaBusiness.contratacion.ComplementosContrato complementosATR_PS; // Pass Subasta
        EndesaBusiness.contratacion.ComplementosContrato complementosATR_PT; // PassThrough

        //EndesaBusiness.facturacion.TamElectricidad tam; 
        EndesaBusiness.medida.TAM tam;

        // Agrupaciones
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_empresa;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_tarifa;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_territorio;

        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_agrupada;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_age;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_agora;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_passthough;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_passpool_periodo;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_passpool_subasta;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_tipo_contrato;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_gestion_atr;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_estado_contrato;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_revendedores;

        Dictionary<DateTime, List<EndesaEntity.facturacion.IRF_Totales_Hist>> dic_totales;


        Dictionary<string, int> dic_maxContrato;

        utilidades.Seguimiento_Procesos ss_pp;

        DateTime fd = new DateTime();
        DateTime fh = new DateTime();

        public InformeRevisionFacturas()
        {
            ss_pp = new Seguimiento_Procesos();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Revision_facturas");
            param = new utilidades.Param("irf_param", servidores.MySQLDB.Esquemas.CON);
            utilidades.Fechas f = new utilidades.Fechas();
            fd = new DateTime(f.UltimoDiaHabil().Year, f.UltimoDiaHabil().Month, 1);
            fh = new DateTime(fd.Year, fd.Month, DateTime.DaysInMonth(fd.Year, fd.Month));

            complementosATR_AGORA = new contratacion.ComplementosContrato(fd, fh, "A01");
            complementosATR_PPH  = new contratacion.ComplementosContrato(fd, fh, "L58");
            complementosATR_PPP = new contratacion.ComplementosContrato(fd, fh, "L59"); 
            complementosATR_PS = new contratacion.ComplementosContrato(fd, fh, "L60");
            complementosATR_PT = new contratacion.ComplementosContrato(fd, fh, "L39");


            dic_agrup_empresa = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_tarifa = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_territorio = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_agrupada = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_age = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_agora = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_passthough = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_passpool_periodo = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_passpool_subasta = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_tipo_contrato = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_gestion_atr = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_estado_contrato = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_revendedores = new Dictionary<string, EndesaEntity.Agrupacion>();
            

        }


        public void Crea_Diaria()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            bool firstOnly = true;
            
            
           
            StringBuilder sb = new StringBuilder();
            string fecha_activacion = "";
            int numreg = 0;

            try
            {

                Crea_Acumulada();
                List<EndesaEntity.facturacion.InformeRevisionFacturas> lista_acumulada
                    = Acumulada();

                strSql = "delete from irf_diaria";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                

                strSql = "SELECT sol.EMP_TIT,"
                    + " sol.CIF, sol.cliente, ps.tipoGestionATR,"
                    + " sol.CUPS_EXT, sol.TARIFA, ps.provincia, ag.tfacagru,"
                    + " sol.codSolATR, sol.estadoSolATR, sol.tipoSolATR, sol.fActivacion,"
                    + " ps.TPUNTMED, t.descripcion as tipo_contrato"
                    + " FROM cont.SOLATRMT sol"
                    + " left outer join cont.PS_AT ps on"
                    + " ps.CUPS22 = sol.CUPS_EXT"
                    + " left outer join cont.irf_cont_fact_agrupada ag on"
                    + " ag.contraext = ps.CONTREXT"
                    + " LEFT OUTER JOIN cont_ticonsps t ON"
                    + " t.tticonps = ps.TTICONPS"
                    + " WHERE" 
                    + " (sol.fActivacion >= " + fd.ToString("yyyyMMdd")
                    + " AND sol.fActivacion <= " + fh.ToString("yyyyMMdd") + ") AND"
                    + " sol.estadoSolATR = 'ACTIVADA' and"
                    + " sol.tipoSolATR IN('BAJA','MODIFICACION','TRASPASO','SUBROGACION')"
                    + " ORDER BY sol.fActivacion desc";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    if (!ExisteEnAcumulada(lista_acumulada,
                        r["CUPS_EXT"].ToString(),
                        Convert.ToDouble(r["codSolATR"]),
                        r["estadoSolATR"].ToString(),
                        r["tipoSolATR"].ToString()))
                    {

                        numreg++;
                        if (firstOnly)
                        {
                            sb.Append("replace into irf_diaria (empresa, cif, cliente, cups22, tarifa, provincia,");
                            sb.Append(" codsolatr, estadosolatr, tiposolatr, factivacion, agora, fact_agrupada,");
                            sb.Append(" tipo_gestion_atr, complemento, tpuntmed, tipo_contrato, f_ult_mod) values ");
                            firstOnly = false;
                        }

                        sb.Append("(").Append(r["EMP_TIT"].ToString()).Append(",");
                        sb.Append("'").Append(r["CIF"].ToString()).Append("',");
                        sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(r["cliente"].ToString())).Append("',");
                        sb.Append("'").Append(r["CUPS_EXT"].ToString()).Append("',");
                        sb.Append("'").Append(r["TARIFA"].ToString()).Append("',");

                        if (r["provincia"] != System.DBNull.Value)
                            sb.Append("'").Append(r["provincia"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        sb.Append(r["codSolATR"].ToString()).Append(",");
                        sb.Append("'").Append(r["estadoSolATR"].ToString()).Append("',");
                        sb.Append("'").Append(r["tipoSolATR"].ToString()).Append("',");

                        fecha_activacion = r["fActivacion"].ToString().Substring(0, 4)
                            + "-" + r["fActivacion"].ToString().Substring(4, 2)
                            + "-" + r["fActivacion"].ToString().Substring(6, 2);

                        sb.Append("'").Append(fecha_activacion).Append("',");

                        if (complementosATR_AGORA.TieneComplemento(r["CUPS_EXT"].ToString().Substring(0, 20)))
                            sb.Append("'").Append("S").Append("',");
                        else
                            sb.Append("'").Append("N").Append("',");

                        if (r["tfacagru"] != System.DBNull.Value)
                        {
                            if(r["tfacagru"].ToString() == "S")
                                sb.Append("'").Append("S").Append("',");
                            else
                                sb.Append("'").Append("N").Append("',");
                        }
                        else
                            sb.Append("null,");

                        if (r["tipoGestionATR"] != System.DBNull.Value)
                            sb.Append(r["tipoGestionATR"].ToString()).Append(",");
                        else
                            sb.Append("null,");


                        if (complementosATR_PPH.TieneComplemento(r["CUPS_EXT"].ToString().Substring(0, 20)))
                            sb.Append("'").Append("PassPool Horario").Append("',");
                        else if(complementosATR_PPP.TieneComplemento(r["CUPS_EXT"].ToString().Substring(0, 20)))
                                sb.Append("'").Append("PassPool Periodo").Append("',");
                        else if (complementosATR_PS.TieneComplemento(r["CUPS_EXT"].ToString().Substring(0, 20)))
                            sb.Append("'").Append("Pass Subasta").Append("',");
                        else
                            sb.Append("null,");


                        if (r["TPUNTMED"] != System.DBNull.Value)
                            sb.Append(Convert.ToInt32(r["TPUNTMED"])).Append(",");
                        else
                            sb.Append("null,");


                        if (r["tipo_contrato"] != System.DBNull.Value)
                            sb.Append("'").Append(r["tipo_contrato"].ToString()).Append("',");
                        else
                            sb.Append("null,");                        

                        sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                        if (numreg == 250)
                        {
                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.CON);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            numreg = 0;
                        }
                    }
                }
                db.CloseConnection();

                if (numreg > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    numreg = 0;
                }

            }
            catch(Exception e)
            {
                ficheroLog.AddError("Crea_Diaria: " + e.Message);
            }


        }

        private void Crea_Acumulada()
        {
            servidores.MySQLDB db;
            MySqlCommand command;            
            string strSql;

            try
            {
                strSql = "replace into irf_acumulada"
                    + " select * from irf_diaria";
                db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Crea_Acumulada: " + e.Message);
            }
        }


        private List<EndesaEntity.facturacion.InformeRevisionFacturas> Diaria()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            List<EndesaEntity.facturacion.InformeRevisionFacturas> l
                = new List<EndesaEntity.facturacion.InformeRevisionFacturas>();

            try
            {

                strSql = "select empresa, cif, cliente, cups22, tarifa, provincia,"
                    + " codsolatr, estadosolatr, tiposolatr, factivacion, tipo_gestion_atr,"
                    + " agora, fact_agrupada, tpuntmed, tipo_contrato, complemento, f_ult_mod " 
                    + " from irf_diaria"
                    + " order by factivacion desc";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.InformeRevisionFacturas c
                        = new EndesaEntity.facturacion.InformeRevisionFacturas();

                    c.empresa = Convert.ToInt32(r["empresa"]);
                    c.cif = r["cif"].ToString();
                    c.cliente = r["cliente"].ToString();
                    c.cups22 = r["cups22"].ToString();
                    c.tarifa = r["tarifa"].ToString();

                    if (r["provincia"] != System.DBNull.Value)
                        c.provincia = r["provincia"].ToString();

                    c.cod_sol_atr = Convert.ToDouble(r["codsolatr"]);
                    c.estado_sol_atr = r["estadosolatr"].ToString();
                    c.tipo_sol_atr = r["tiposolatr"].ToString();
                    c.fecha_activacion = Convert.ToDateTime(r["factivacion"]);
                    c.agora = r["agora"].ToString() == "S";
                    c.agrupada = r["fact_agrupada"].ToString() == "S";
                    c.fecha_anexion = Convert.ToDateTime(r["f_ult_mod"]);

                    if (r["tpuntmed"] != System.DBNull.Value)
                        c.tpuntmed = Convert.ToInt32(r["tpuntmed"]);

                    if (r["tipo_gestion_atr"] != System.DBNull.Value)
                        c.tipo_gestion_atr = Convert.ToInt32(r["tipo_gestion_atr"]);

                    if (r["tipo_contrato"] != System.DBNull.Value)
                        c.tipo_contrato = r["tipo_contrato"].ToString();

                    if (r["complemento"] != System.DBNull.Value)
                        c.complemento = r["complemento"].ToString();

                    l.Add(c);

                }
                db.CloseConnection();
                return l;

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Crea_Acumulada: " + e.Message);
                return null;
            }
        }

        private List<EndesaEntity.facturacion.InformeRevisionFacturas> Acumulada()
        {
            DateTime fd = new DateTime();
            DateTime fh = new DateTime();
            utilidades.Fechas f = new utilidades.Fechas();

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            List<EndesaEntity.facturacion.InformeRevisionFacturas> l
                = new List<EndesaEntity.facturacion.InformeRevisionFacturas>();

            try
            {
                fd = new DateTime(f.UltimoDiaHabil().Year, f.UltimoDiaHabil().Month, 1);
                fh = new DateTime(fd.Year, fd.Month, DateTime.DaysInMonth(fd.Year, fd.Month));

                strSql = "select empresa, cif, cliente, cups22, tarifa, provincia,"
                    + " codsolatr, estadosolatr, tiposolatr, factivacion, complemento,"
                    + " agora, fact_agrupada, tpuntmed, tipo_contrato, tipo_gestion_atr,"
                    + " f_ult_mod"                    
                    + " from irf_acumulada where"
                    + " (fActivacion >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " AND fActivacion <= '" + fh.ToString("yyyy-MM-dd") + "')"
                    + " order by factivacion desc";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    EndesaEntity.facturacion.InformeRevisionFacturas c
                        = new EndesaEntity.facturacion.InformeRevisionFacturas();

                    c.empresa = Convert.ToInt32(r["empresa"]);
                    c.cif = r["cif"].ToString();
                    c.cliente = r["cliente"].ToString();
                    c.cups22 = r["cups22"].ToString();
                    c.tarifa = r["tarifa"].ToString();

                    if (r["tpuntmed"] != System.DBNull.Value)
                        c.provincia = r["provincia"].ToString();

                    c.cod_sol_atr = Convert.ToDouble(r["codsolatr"]);
                    c.tipo_sol_atr = r["tiposolatr"].ToString();
                    c.estado_sol_atr = r["estadosolatr"].ToString();
                    c.fecha_activacion = Convert.ToDateTime(r["factivacion"]);
                    c.agora = r["agora"].ToString() == "S";
                    c.agrupada = r["fact_agrupada"].ToString() == "S";
                    c.fecha_anexion = Convert.ToDateTime(r["f_ult_mod"]);

                    if (r["tpuntmed"] != System.DBNull.Value)
                        c.tpuntmed = Convert.ToInt32(r["tpuntmed"]);

                    if (r["tipo_gestion_atr"] != System.DBNull.Value)
                        c.tipo_gestion_atr = Convert.ToInt32(r["tipo_gestion_atr"]);

                    if (r["tipo_contrato"] != System.DBNull.Value)
                        c.tipo_contrato = r["tipo_contrato"].ToString();

                    if (r["complemento"] != System.DBNull.Value)
                        c.complemento = r["complemento"].ToString();

                    l.Add(c);

                }
                db.CloseConnection();
                return l;

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Crea_Acumulada: " + e.Message);
                return null;
            }
        }


        public void GeneraInformeAgora()
        {
            int c = 1;
            int f = 1;
            string ruta_salida_archivo = "";
            utilidades.Fechas utilFechas = new utilidades.Fechas();
            DateTime ultimo_dia_habil = new DateTime();

            try
            {

                ultimo_dia_habil = utilFechas.UltimoDiaHabil();

                List<EndesaEntity.facturacion.InformeRevisionFacturas> lista_diaria
                    = Diaria();
                List<EndesaEntity.facturacion.InformeRevisionFacturas> lista_acumulada
                    = Acumulada();

                EndesaBusiness.utilidades.Fichero.BorrarArchivos_MenosUltimo(param.GetValue("Ubicacion_Informes"),
                    param.GetValue("Excel_prefijo_Agora"), "xlsx", null);

                ruta_salida_archivo = param.GetValue("Ubicacion_Informes")
                    + param.GetValue("Excel_prefijo_Agora")
                    + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";

                FileInfo fileInfo = new FileInfo(ruta_salida_archivo);

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("DIARIA");
                var headerCells = workSheet.Cells[1, 1, 1, 13];
                var headerFont = headerCells.Style.Font;

                workSheet.View.FreezePanes(2, 1);

                headerFont.Bold = true;
                workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                workSheet.Cells[f, c].Value = "CIF"; c++;
                workSheet.Cells[f, c].Value = "RAZÓN SOCIAL"; c++;
                workSheet.Cells[f, c].Value = "CUPS22"; c++;
                workSheet.Cells[f, c].Value = "TARIFA"; c++;
                workSheet.Cells[f, c].Value = "PROVINCIA"; c++;
                workSheet.Cells[f, c].Value = "ESTADO SOLICITUD"; c++;
                workSheet.Cells[f, c].Value = "TIPO SOLICITUD"; c++;
                workSheet.Cells[f, c].Value = "FECHA ACTIVACIÓN"; c++;
                workSheet.Cells[f, c].Value = "TIPO PUNTO MEDIDA"; c++;
                workSheet.Cells[f, c].Value = "FACT. AGRUPADA"; c++;
                workSheet.Cells[f, c].Value = "TIPO GESTION ATR"; c++;
                workSheet.Cells[f, c].Value = "TIPO CONTRATO";
                

                for (int i = 1; i <= c; i++)
                {
                    workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }


                foreach (EndesaEntity.facturacion.InformeRevisionFacturas p in lista_diaria)
                {
                    if (p.agora)
                    {
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = p.empresa; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.cliente; c++;
                        workSheet.Cells[f, c].Value = p.cups22; c++;
                        workSheet.Cells[f, c].Value = p.tarifa; c++;
                        workSheet.Cells[f, c].Value = p.provincia; c++;
                        workSheet.Cells[f, c].Value = p.estado_sol_atr; c++;
                        workSheet.Cells[f, c].Value = p.tipo_sol_atr; c++;
                        workSheet.Cells[f, c].Value = p.fecha_activacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                        if(p.tpuntmed > 0)
                            workSheet.Cells[f, c].Value = p.tpuntmed;


                        c++;
                        workSheet.Cells[f, c].Value = (p.agrupada ? "S" : "N");
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                        if(p.tipo_gestion_atr > 0)
                            workSheet.Cells[f, c].Value = p.tipo_gestion_atr;

                        c++;

                        workSheet.Cells[f, c].Value = p.tipo_contrato;

                        if (p.fecha_activacion.Date != ultimo_dia_habil.Date)
                            for (int i = 1; i <= c; i++)
                            {
                                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                                    System.Drawing.Color.FromArgb(255, 100, 100));

                            }

                        if (p.empresa == 70)
                            for (int i = 1; i <= c; i++)
                            {
                                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                                    System.Drawing.Color.FromArgb(0, 128, 255));

                            }

                        if (p.tipo_contrato == "DE TEMPORADA" || p.tipo_contrato == "SOCORRO" || p.tipo_contrato == "EVENTUAL")
                            for (int i = 1; i <= c; i++)
                            {
                                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                                    System.Drawing.Color.FromArgb(255, 255, 128));

                            }


                    }


                }

                var allCells = workSheet.Cells[1, 1, f, c];
                workSheet.Cells["A1:M1"].AutoFilter = true;


                headerFont.Bold = true;
                allCells.AutoFitColumns();


                workSheet = excelPackage.Workbook.Worksheets.Add("ACUMULADA");
                headerFont.Bold = true;

                f = 1;
                c = 1;

                headerFont.Bold = true;
                workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                workSheet.Cells[f, c].Value = "CIF"; c++;
                workSheet.Cells[f, c].Value = "RAZÓN SOCIAL"; c++;
                workSheet.Cells[f, c].Value = "CUPS22"; c++;
                workSheet.Cells[f, c].Value = "TARIFA"; c++;
                workSheet.Cells[f, c].Value = "PROVINCIA"; c++;
                workSheet.Cells[f, c].Value = "ESTADO SOLICITUD"; c++;
                workSheet.Cells[f, c].Value = "TIPO SOLICITUD"; c++;
                workSheet.Cells[f, c].Value = "FECHA ACTIVACIÓN"; c++;
                workSheet.Cells[f, c].Value = "TIPO PUNTO MEDIDA"; c++;
                workSheet.Cells[f, c].Value = "TIPO CONTRATO"; c++;
                workSheet.Cells[f, c].Value = "FACT. AGRUPADA"; c++;
                workSheet.Cells[f, c].Value = "TIPO GESTION ATR"; c++;
                workSheet.Cells[f, c].Value = "FECHA ANEXIÓN"; 

                for (int i = 1; i <= c; i++)
                {
                    workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                foreach (EndesaEntity.facturacion.InformeRevisionFacturas p in lista_diaria)
                {
                    if (p.agora)
                    {
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = p.empresa; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.cliente; c++;
                        workSheet.Cells[f, c].Value = p.cups22; c++;
                        workSheet.Cells[f, c].Value = p.tarifa; c++;
                        workSheet.Cells[f, c].Value = p.provincia; c++;
                        workSheet.Cells[f, c].Value = p.estado_sol_atr; c++;
                        workSheet.Cells[f, c].Value = p.tipo_sol_atr; c++;
                        workSheet.Cells[f, c].Value = p.fecha_activacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                        if (p.tpuntmed > 0)
                            workSheet.Cells[f, c].Value = p.tpuntmed;

                        c++;
                        workSheet.Cells[f, c].Value = p.tipo_contrato; c++;

                        workSheet.Cells[f, c].Value = (p.agrupada ? "S" : "N");
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                        if (p.tipo_gestion_atr > 0)
                            workSheet.Cells[f, c].Value = p.tipo_gestion_atr;

                        c++;

                        workSheet.Cells[f, c].Value = p.fecha_anexion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;


                        //if (p.empresa == 70)
                        //    for (int i = 1; i <= c; i++)
                        //    {
                        //        workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        //        workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                        //            System.Drawing.Color.FromArgb(0, 128, 255));

                        //    }

                        if (p.tipo_contrato == "DE TEMPORADA" || p.tipo_contrato == "SOCORRO" || p.tipo_contrato == "EVENTUAL")
                            for (int i = 1; i <= c; i++)
                            {
                                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                                    System.Drawing.Color.FromArgb(255, 255, 128));

                            }


                    }


                }


                foreach (EndesaEntity.facturacion.InformeRevisionFacturas p in lista_acumulada)
                {
                    if (p.agora)
                    {
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = p.empresa; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.cliente; c++;
                        workSheet.Cells[f, c].Value = p.cups22; c++;
                        workSheet.Cells[f, c].Value = p.tarifa; c++;
                        workSheet.Cells[f, c].Value = p.provincia; c++;
                        workSheet.Cells[f, c].Value = p.estado_sol_atr; c++;
                        workSheet.Cells[f, c].Value = p.tipo_sol_atr; c++;
                        workSheet.Cells[f, c].Value = p.fecha_activacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                        if (p.tpuntmed > 0)
                            workSheet.Cells[f, c].Value = p.tpuntmed;
                        c++;

                        workSheet.Cells[f, c].Value = p.tipo_contrato;c++;

                        workSheet.Cells[f, c].Value = (p.agrupada ? "S" : "N");
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                        if (p.tipo_gestion_atr > 0)
                            workSheet.Cells[f, c].Value = p.tipo_gestion_atr;
                        c++;

                        workSheet.Cells[f, c].Value = p.fecha_anexion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                                                      

                        //if (p.empresa == 70)
                        //    for (int i = 1; i <= c; i++)
                        //    {
                        //        workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        //        workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                        //            System.Drawing.Color.FromArgb(0, 128, 255));

                        //    }

                        if (p.tipo_contrato == "DE TEMPORADA" || p.tipo_contrato == "SOCORRO" || p.tipo_contrato == "EVENTUAL")
                            for (int i = 1; i <= c; i++)
                            {
                                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                                    System.Drawing.Color.FromArgb(255, 255, 128));

                            }

                    }

                }

                headerCells = workSheet.Cells[1, 1, 1, c];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;
                allCells = workSheet.Cells[1, 1, f, c];

                allCells.AutoFitColumns();

                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:M1"].AutoFilter = true;
                allCells.AutoFitColumns();

                excelPackage.Save();

                EnvioCorreoAgora(ruta_salida_archivo);

            }
            catch(Exception e)
            {
                ficheroLog.AddError("GeneraInformeAgora: " + e.Message);
            }
        }
        public void GeneraInformeNoAgora()
        {
            int c = 1;
            int f = 1;
            string ruta_salida_archivo = "";

            utilidades.Fechas utilFechas = new utilidades.Fechas();
            DateTime ultimo_dia_habil = new DateTime();

            try
            {

                ultimo_dia_habil = utilFechas.UltimoDiaHabil();

                List<EndesaEntity.facturacion.InformeRevisionFacturas> lista_diaria
                    = Diaria();
                List<EndesaEntity.facturacion.InformeRevisionFacturas> lista_acumulada
                    = Acumulada();


                EndesaBusiness.utilidades.Fichero.BorrarArchivos_MenosUltimo(param.GetValue("Ubicacion_Informes"), 
                    param.GetValue("Excel_prefijo_NoAgora"), "xlsx", null);

                ruta_salida_archivo = param.GetValue("Ubicacion_Informes")
                   + param.GetValue("Excel_prefijo_NoAgora")
                   + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";                

                FileInfo fileInfo = new FileInfo(ruta_salida_archivo);

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("DIARIA");
                var headerCells = workSheet.Cells[1, 1, 1, 14];
                var headerFont = headerCells.Style.Font;

                workSheet.View.FreezePanes(2, 1);

                headerFont.Bold = true;
                workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                workSheet.Cells[f, c].Value = "CIF"; c++;
                workSheet.Cells[f, c].Value = "RAZÓN SOCIAL"; c++;
                workSheet.Cells[f, c].Value = "CUPS22"; c++;
                workSheet.Cells[f, c].Value = "TARIFA"; c++;
                workSheet.Cells[f, c].Value = "PROVINCIA"; c++;
                workSheet.Cells[f, c].Value = "ESTADO SOLICITUD"; c++;
                workSheet.Cells[f, c].Value = "TIPO SOLICITUD"; c++;
                workSheet.Cells[f, c].Value = "FECHA ACTIVACIÓN"; c++;
                workSheet.Cells[f, c].Value = "TIPO PUNTO MEDIDA"; c++;
                workSheet.Cells[f, c].Value = "FACT. AGRUPADA"; c++;
                workSheet.Cells[f, c].Value = "TIPO GESTION ATR"; c++;
                workSheet.Cells[f, c].Value = "COMPLEMENTO"; c++;
                workSheet.Cells[f, c].Value = "TIPO CONTRATO"; 
                

                for (int i = 1; i <= c; i++)
                {
                    workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }


                foreach (EndesaEntity.facturacion.InformeRevisionFacturas p in lista_diaria)
                {
                    if (!p.agora)
                    {
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = p.empresa; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.cliente; c++;
                        workSheet.Cells[f, c].Value = p.cups22; c++;
                        workSheet.Cells[f, c].Value = p.tarifa; c++;
                        workSheet.Cells[f, c].Value = p.provincia; c++;
                        workSheet.Cells[f, c].Value = p.estado_sol_atr; c++;
                        workSheet.Cells[f, c].Value = p.tipo_sol_atr; c++;
                        workSheet.Cells[f, c].Value = p.fecha_activacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                        if (p.tpuntmed > 0)
                            workSheet.Cells[f, c].Value = p.tpuntmed;
                        c++;

                        workSheet.Cells[f, c].Value = (p.agrupada ? "S" : "N"); c++;

                        if (p.tipo_gestion_atr > 0)
                            workSheet.Cells[f, c].Value = p.tipo_gestion_atr;
                        c++;

                        workSheet.Cells[f, c].Value = p.complemento; c++;

                        workSheet.Cells[f, c].Value = p.tipo_contrato;
                        

                        if (p.fecha_activacion.Date != ultimo_dia_habil.Date)                        
                            for (int i = 1; i <= c; i++)
                            {                                
                                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                                    System.Drawing.Color.FromArgb(255, 100, 100));

                            }

                        //if (p.empresa == 70)
                        //    for (int i = 1; i <= c; i++)
                        //    {
                        //        workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        //        workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                        //            System.Drawing.Color.FromArgb(0, 128, 255));

                        //    }

                        if (p.tipo_contrato == "DE TEMPORADA" || p.tipo_contrato == "SOCORRO" || p.tipo_contrato == "EVENTUAL")
                            for (int i = 1; i <= c; i++)
                            {
                                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                                    System.Drawing.Color.FromArgb(255, 255, 128));

                            }


                    }


                }

                var allCells = workSheet.Cells[1, 1, f, c];
                workSheet.Cells["A1:N1"].AutoFilter = true;


                headerFont.Bold = true;
                allCells.AutoFitColumns();


                workSheet = excelPackage.Workbook.Worksheets.Add("ACUMULADA");
                headerFont.Bold = true;

                f = 1;
                c = 1;

                headerFont.Bold = true;
                workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                workSheet.Cells[f, c].Value = "CIF"; c++;
                workSheet.Cells[f, c].Value = "RAZÓN SOCIAL"; c++;
                workSheet.Cells[f, c].Value = "CUPS22"; c++;
                workSheet.Cells[f, c].Value = "TARIFA"; c++;
                workSheet.Cells[f, c].Value = "PROVINCIA"; c++;
                workSheet.Cells[f, c].Value = "ESTADO SOLICITUD"; c++;
                workSheet.Cells[f, c].Value = "TIPO SOLICITUD"; c++;
                workSheet.Cells[f, c].Value = "FECHA ACTIVACIÓN"; c++;
                workSheet.Cells[f, c].Value = "TIPO PUNTO MEDIDA"; c++;
                workSheet.Cells[f, c].Value = "TIPO CONTRATO"; c++;
                workSheet.Cells[f, c].Value = "FACT. AGRUPADA"; c++;
                workSheet.Cells[f, c].Value = "TIPO GESTION ATR"; c++;
                workSheet.Cells[f, c].Value = "COMPLEMENTO"; c++;
                workSheet.Cells[f, c].Value = "FECHA ANEXIÓN";

                for (int i = 1; i <= c; i++)
                {
                    workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                foreach (EndesaEntity.facturacion.InformeRevisionFacturas p in lista_diaria)
                {
                    if (!p.agora)
                    {
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = p.empresa; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.cliente; c++;
                        workSheet.Cells[f, c].Value = p.cups22; c++;
                        workSheet.Cells[f, c].Value = p.tarifa; c++;
                        workSheet.Cells[f, c].Value = p.provincia; c++;                        
                        workSheet.Cells[f, c].Value = p.estado_sol_atr; c++;
                        workSheet.Cells[f, c].Value = p.tipo_sol_atr; c++;
                        workSheet.Cells[f, c].Value = p.fecha_activacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                                               

                        if (p.tpuntmed > 0)
                            workSheet.Cells[f, c].Value = p.tpuntmed;
                        c++;

                        workSheet.Cells[f, c].Value = p.tipo_contrato;c++;
                        workSheet.Cells[f, c].Value = (p.agrupada ? "S" : "N");
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                        if (p.tipo_gestion_atr > 0)
                            workSheet.Cells[f, c].Value = p.tipo_gestion_atr;

                        c++;

                        workSheet.Cells[f, c].Value = p.complemento; c++;

                        workSheet.Cells[f, c].Value = p.fecha_anexion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                              

                        //if (p.empresa == 70)
                        //    for (int i = 1; i <= c; i++)
                        //    {
                        //        workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        //        workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                        //            System.Drawing.Color.FromArgb(0, 128, 255));

                        //    }

                        if (p.tipo_contrato == "DE TEMPORADA" || p.tipo_contrato == "SOCORRO" || p.tipo_contrato == "EVENTUAL")
                            for (int i = 1; i <= c; i++)
                            {
                                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                                    System.Drawing.Color.FromArgb(255, 255, 128));

                            }


                    }


                }



                foreach (EndesaEntity.facturacion.InformeRevisionFacturas p in lista_acumulada)
                {
                    if (!p.agora)
                    {
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = p.empresa; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.cliente; c++;
                        workSheet.Cells[f, c].Value = p.cups22; c++;
                        workSheet.Cells[f, c].Value = p.tarifa; c++;
                        workSheet.Cells[f, c].Value = p.provincia; c++;
                        workSheet.Cells[f, c].Value = p.estado_sol_atr; c++;
                        workSheet.Cells[f, c].Value = p.tipo_sol_atr; c++;
                        workSheet.Cells[f, c].Value = p.fecha_activacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                        if (p.tpuntmed > 0)
                            workSheet.Cells[f, c].Value = p.tpuntmed;

                        c++;

                        workSheet.Cells[f, c].Value = p.tipo_contrato; c++;

                        workSheet.Cells[f, c].Value = (p.agrupada ? "S" : "N");
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                        if (p.tipo_gestion_atr > 0)
                            workSheet.Cells[f, c].Value = p.tipo_gestion_atr;

                        c++;

                        workSheet.Cells[f, c].Value = p.complemento; c++;

                        workSheet.Cells[f, c].Value = p.fecha_anexion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                        if (p.empresa == 70)
                            for (int i = 1; i <= c; i++)
                            {
                                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                                    System.Drawing.Color.FromArgb(0, 128, 255));

                            }

                        if (p.tipo_contrato == "DE TEMPORADA" || p.tipo_contrato == "SOCORRO" || p.tipo_contrato == "EVENTUAL")
                            for (int i = 1; i <= c; i++)
                            {
                                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                                    System.Drawing.Color.FromArgb(255, 255, 128));

                            }

                    }

                }

                headerCells = workSheet.Cells[1, 1, 1, c];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;
                allCells = workSheet.Cells[1, 1, f, c];

                allCells.AutoFitColumns();

                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:O1"].AutoFilter = true;
                allCells.AutoFitColumns();

                excelPackage.Save();

                EnvioCorreoNoAgora(ruta_salida_archivo);


            }
            catch (Exception e)
            {
                ficheroLog.AddError("GeneraInformeNoAgora: " + e.Message);
            }
        }
        public void Inventario_por_Tipologias()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            bool firstOnly = true;
            int f = 1;
            int c = 1;
            EndesaEntity.Agrupacion o;

            // Totales para agrupaciones

            int total_agrupada = 0;
            int total_age = 0;
            int total_agora = 0;
            int total_passthough = 0;
            int total_passpool_horario = 0;
            int total_passpool_periodo = 0;
            int total_passpool_subasta = 0;
            int total_gestion_atr = 0;
            int total_revendedores = 0;
            int total_medido_baja = 0;
            int total_multipuntos = 0;
            double valor_tam = 0;
            int total_reg = 0;

            int total_catalogo_mt = 0;
            int total_personalizado_r_esp = 0;
            int total_personalizado_r_est = 0;
            int total_flexible = 0;
                

            medida.Perdidas perdidas = new medida.Perdidas();
            medida.SCEA scea = new medida.SCEA();

            utilidades.Fechas utilfecha = new Fechas();

            contratacion.ContratosPS_Portugal_MT complementos_PS =
                new contratacion.ContratosPS_Portugal_MT();
            complementos_PS.CargaInventarioMT_ESP();

            FileInfo file;

            try
            {
                if (ss_pp.GetFecha_FinProceso("Facturación", "Informe Inventario Tipologias ES", "Informe Inventario Tipologias ES").Date
                    < ss_pp.GetFecha_FinProceso("Facturación", "Informe Pendiente BI", "Informe Pendiente BI").Date)
                {
                    ss_pp.Update_Fecha_Inicio("Facturación", "Informe Inventario Tipologias ES", "Informe Inventario Tipologias ES");

                    string[] listaArchivos = System.IO.Directory.GetFiles(param.GetValue("Ubicacion_Informes"),
                       param.GetValue("Excel_prefijo_Tipologias") + "*.xlsx");

                    for (int i = 0; i < listaArchivos.Length; i++)
                    {
                        file = new FileInfo(listaArchivos[i]);
                        file.Delete();
                    }



                    Console.WriteLine("Inventario_por_Tipologias");
                    Console.WriteLine("=========================");


                    Console.WriteLine("Cargamos tabla PS_AT");
                    EndesaBusiness.contratacion.PS_AT pS_AT = new contratacion.PS_AT();

                    Console.WriteLine("Cargamos TAM para los puntos de PS_AT");
                    //tam = new TamElectricidad(pS_AT.dic.Select(z => z.Value.cups20).ToList());
                    //tam = new medida.TAM();
                    //tam.CargaTAM();


                    Console.WriteLine("Cargamos SCEA Multipuntos");
                    scea.CargaMultipuntos();


                    dic_maxContrato = CargaUltimaVersion();

                    EndesaBusiness.utilidades.Fichero.BorrarArchivos_MenosUltimo(param.GetValue("Ubicacion_Informes"),
                        param.GetValue("Excel_prefijo_Tipologias"), "xlsx", null);

                    string ruta_salida_archivo = param.GetValue("Ubicacion_Informes")
                      + param.GetValue("Excel_prefijo_Tipologias")
                      + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";

                    Console.WriteLine("Generando archivo: " + ruta_salida_archivo);

                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                    FileInfo fileInfo = new FileInfo(ruta_salida_archivo);
                    ExcelPackage excelPackage = new ExcelPackage(fileInfo);

                    var workSheet = excelPackage.Workbook.Worksheets.Add("PS");
                    var headerCells = workSheet.Cells[1, 1, 1, 25];
                    var headerFont = headerCells.Style.Font;

                    workSheet.View.FreezePanes(2, 1);

                    headerFont.Bold = true;

                    //strSql = "SELECT ps.IDU,"
                    //    + " if(ps.CUPS22 IS NULL OR ps.CUPS22 = '','XXXXXXXXXXXXXXXXXXXXXX',ps.CUPS22) AS CUPS22,"
                    //     + " ps.EMPRESA, ps.TARIFA,"
                    //     + " ps.NIF , ps.Cliente,"
                    //     + " ps.provincia,  t.Territorio, ps.CONTREXT, ps.TPUNTMED,"
                    //     + " if (ps.TTICONPS = 'E', 'Si','No') AS eventual,"
                    //     + " if (ps.tipoGestionATR = 1, 'No', 'Si') AS gestionpropia_atr,"
                    //     + " ag.tfacagru, tt.descripcion as tipo_contrato,"
                    //     + " if(age.CUPS22 is null, 'No', 'Si') as Producto, ttt.TAM,"
                    //     + " e.Descripcion AS estado_contrato, ps.Version, ps.FPSERCON,"
                    //     + " if(s.tdistri = 'REV', 'Si','No') AS revendedor"                     
                    //     + " FROM cont.PS_AT ps"
                    //     + " INNER JOIN cont.paramEmpresaTitular eett ON"
                    //     + " eett.empresaTitular = ps.EMPRESA"
                    //     + " LEFT OUTER JOIN fact.t_territorios t ON"
                    //     + " t.Provincia = ps.provincia"
                    //     + " left outer join cont.irf_cont_fact_agrupada ag on"
                    //     + " ag.contraext = ps.CONTREXT"
                    //     + " LEFT OUTER JOIN cont.cont_ticonsps tt ON"
                    //     + " tt.tticonps = ps.TTICONPS"
                    //     + " LEFT OUTER JOIN cont.irf_age age on"
                    //     + " substr(age.CUPS22,1,20) = substr(ps.CUPS22,1,20)"
                    //     + " LEFT OUTER JOIN fact.tam ttt on"
                    //     + " ttt.CEMPTITU = eett.codEmp and"
                    //     + " ttt.CNIFDNIC = ps.NIF and"
                    //     + " ttt.CUPS20 = substr(ps.CUPS22,1,20)"
                    //     + " LEFT OUTER JOIN cont.cont_estadoscontrato e on"
                    //     + " e.Cod_Estado = ps.estadoCont"
                    //     + " LEFT OUTER JOIN med.scea s ON"
                    //     + " s.IDU = ps.IDU";

                    strSql = "SELECT ps.IDU,"
                        + " if(ps.CUPS22 IS NULL OR ps.CUPS22 = '','XXXXXXXXXXXXXXXXXXXXXX',ps.CUPS22) AS CUPS22,"
                         + " ps.EMPRESA, ps.TARIFA,"
                         + " ps.NIF , ps.Cliente,"
                         + " ps.provincia,  t.Territorio, ps.CONTREXT, ps.TPUNTMED,"
                         + " if (ps.TTICONPS = 'E', 'Si','No') AS eventual,"
                         + " if (ps.tipoGestionATR = 1, 'No', 'Si') AS gestionpropia_atr,"
                         + " ag.tfacagru, tt.descripcion as tipo_contrato,"
                         //+ " if(age.CCOUNIPS is null, 'No', 'Si') as Producto, ttt.TAM,"
                         + " ttt.TAM,"
                         + " if(age.CCOUNIPS is null, 'No', 'Si') as Producto,"
                         + " e.Descripcion AS estado_contrato, ps.Version, ps.FPSERCON,"
                         + " if(s.tdistri = 'REV', 'Si','No') AS revendedor,"
                         + " agr.fh_periodo AS PERIODO_PENDIENTE, agr.cd_estado, de.de_estado ,agr.cd_subestado, ds.de_subestado"
                         + " FROM cont.PS_AT ps"
                         + " INNER JOIN cont.paramEmpresaTitular eett ON"
                         + " eett.empresaTitular = ps.EMPRESA"
                         + " LEFT OUTER JOIN fact.t_territorios t ON"
                         + " t.Provincia = ps.provincia"
                         + " left outer join cont.irf_cont_fact_agrupada ag on"
                         + " ag.contraext = ps.CONTREXT"
                         + " LEFT OUTER JOIN cont.cont_ticonsps tt ON"
                         + " tt.tticonps = ps.TTICONPS"
                         + " LEFT OUTER JOIN cont.irf_s20vge age on"
                         + " age.CCOUNIPS = ps.IDU"
                         + " LEFT OUTER JOIN fact.tam ttt on"
                         + " ttt.CEMPTITU = eett.codEmp and"
                         + " ttt.CNIFDNIC = ps.NIF and"
                         + " ttt.CUPS20 = substr(ps.CUPS22,1,20)"
                         + " LEFT OUTER JOIN cont.cont_estadoscontrato e on"
                         + " e.Cod_Estado = ps.estadoCont"
                         + " LEFT OUTER JOIN med.scea s ON"
                         + " s.IDU = ps.IDU"
                         + " LEFT OUTER JOIN fact.t_ed_h_sap_pendiente_facturar_agrupado agr on"
                         + " agr.cd_cups = substr(ps.CUPS22,1,20) AND agr.fh_envio = (SELECT MAX(fh_envio) FROM fact.t_ed_h_sap_pendiente_facturar_agrupado)"
                         + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de ON"
                         + " de.cd_estado = agr.cd_estado"
                         + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds ON"
                         + " ds.cd_subestado = agr.cd_subestado";


                    Console.WriteLine(strSql);
                    ficheroLog.Add(strSql);

                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        total_reg++;

                        Console.CursorLeft = 0;
                        Console.Write("Exportando: " + total_reg.ToString("N0"));

                        c = 1;
                        #region Cabecera
                        if (firstOnly)
                        {

                            workSheet.Cells[f, c].Value = "CUPS"; c++;
                            workSheet.Cells[f, c].Value = "CUPS22"; c++;
                            workSheet.Cells[f, c].Value = "CIF"; c++;
                            workSheet.Cells[f, c].Value = "RAZÓN SOCIAL"; c++;
                            workSheet.Cells[f, c].Value = "CONVERTIDO"; c++;
                            workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                            workSheet.Cells[f, c].Value = "TARIFA"; c++;
                            workSheet.Cells[f, c].Value = "TIPO PUNTO SUMINISTRO"; c++;
                            workSheet.Cells[f, c].Value = "PROVINCIA"; c++;
                            workSheet.Cells[f, c].Value = "TERRITORIO"; c++;
                            workSheet.Cells[f, c].Value = "CONTRATO EXT"; c++;
                            workSheet.Cells[f, c].Value = "AGRUPADA"; c++;
                            workSheet.Cells[f, c].Value = "AGE";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;
                            workSheet.Cells[f, c].Value = "AGORA";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;
                            workSheet.Cells[f, c].Value = "PASSTHOUGH"; c++;
                            workSheet.Cells[f, c].Value = "PASSPOOL HORARIO"; c++;
                            workSheet.Cells[f, c].Value = "PASSPOOL PERIODO"; c++;
                            workSheet.Cells[f, c].Value = "PASSPOOL SUBASTA"; c++;
                            workSheet.Cells[f, c].Value = "TIPO CONTRATO"; c++;
                            workSheet.Cells[f, c].Value = "GESTIÓN PROPIA ATR"; c++;
                            workSheet.Cells[f, c].Value = "TAM";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;
                            workSheet.Cells[f, c].Value = "ESTADO CONTRATO"; c++;
                            workSheet.Cells[f, c].Value = "VERSIÓN"; c++;
                            workSheet.Cells[f, c].Value = "F. PUESTA SERVICIO"; c++;
                            workSheet.Cells[f, c].Value = "REVENDEDORES"; c++;
                            workSheet.Cells[f, c].Value = "ÚLTIMO CONVERTIDO"; c++;
                            workSheet.Cells[f, c].Value = "MEDIDO EN BAJA"; c++;
                            workSheet.Cells[f, c].Value = "MULTIPUNTO"; c++;
                            workSheet.Cells[f, c].Value = "CATÁLOGO"; c++;
                            workSheet.Cells[f, c].Value = "PERIODO PENDIENTE"; c++;
                            workSheet.Cells[f, c].Value = "ESTADO PENDIENTE"; c++;
                            workSheet.Cells[f, c].Value = "SUBESTADO PENDIENTE"; c++;

                            firstOnly = false;
                        }
                        c = 1;
                        f++;

                        #endregion

                        if (r["IDU"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["IDU"].ToString();
                        c++;

                        if (r["CUPS22"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CUPS22"].ToString();
                        c++;

                        if (r["NIF"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["NIF"].ToString();
                        c++;

                        if (r["Cliente"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["Cliente"].ToString();
                        c++;

                        if (r["TARIFA"] != System.DBNull.Value)
                        {
                            if (r["TARIFA"].ToString().Contains("TD"))
                            {
                                workSheet.Cells[f, c].Value = "Si";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            else
                            {
                                workSheet.Cells[f, c].Value = "No";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }

                        }
                        else
                        {
                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        c++;


                        if (r["EMPRESA"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["EMPRESA"].ToString();

                            if (!dic_agrup_empresa.TryGetValue(r["EMPRESA"].ToString(), out o))
                            {
                                o = new EndesaEntity.Agrupacion();
                                o.tipo = r["EMPRESA"].ToString();
                                o.total = 1;
                                dic_agrup_empresa.Add(o.tipo, o);
                            }
                            else
                                o.total++;


                        }

                        c++;

                        if (r["TARIFA"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                            if (!dic_agrup_tarifa.TryGetValue(r["TARIFA"].ToString(), out o))
                            {
                                o = new EndesaEntity.Agrupacion();
                                o.tipo = r["TARIFA"].ToString();
                                o.total = 1;
                                dic_agrup_tarifa.Add(o.tipo, o);
                            }
                            else
                                o.total++;

                        }

                        c++;

                        if (r["TPUNTMED"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["TPUNTMED"].ToString();
                        }

                        c++;

                        if (r["provincia"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["provincia"].ToString();
                        c++;

                        if (r["Territorio"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["Territorio"].ToString();
                            if (!dic_agrup_territorio.TryGetValue(r["Territorio"].ToString(), out o))
                            {
                                o = new EndesaEntity.Agrupacion();
                                o.tipo = r["Territorio"].ToString();
                                o.total = 1;
                                dic_agrup_territorio.Add(o.tipo, o);
                            }
                            else
                                o.total++;

                        }

                        c++;

                        if (r["CONTREXT"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CONTREXT"].ToString();
                        c++;

                        if (r["tfacagru"] != System.DBNull.Value)
                        {

                            if (r["tfacagru"].ToString() == "S")
                            {

                                total_agrupada++;
                                workSheet.Cells[f, c].Value = "Si";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            else
                            {

                                workSheet.Cells[f, c].Value = "No";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }

                        }
                        else
                        {
                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        c++;


                        if (r["Producto"] != System.DBNull.Value)
                        {
                            if (r["Producto"].ToString() == "Si")
                                total_age++;

                            workSheet.Cells[f, c].Value = r["Producto"].ToString();
                        }

                        // AGE Adminitraciones Publicas

                        //if (r["NIF"] != System.DBNull.Value)
                        //{
                        //    if(r["NIF"].ToString().Substring(0,1) == "P" ||
                        //        r["NIF"].ToString().Substring(0, 1) == "Q" ||
                        //        r["NIF"].ToString().Substring(0, 1) == "S")
                        //    {
                        //        total_age++;
                        //        workSheet.Cells[f, c].Value = "Si";
                        //        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //    }
                        //    else
                        //    {
                        //        workSheet.Cells[f, c].Value = "No";
                        //        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //    }
                        //}

                        c++;


                        if (complementosATR_AGORA.TieneComplemento(r["CUPS22"].ToString().Substring(0, 20)))
                        {

                            total_agora++;
                            workSheet.Cells[f, c].Value = "Si";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else
                        {

                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        c++;

                        if (complementosATR_PT.TieneComplemento(r["CUPS22"].ToString().Substring(0, 20)))
                        {
                            total_passthough++;
                            workSheet.Cells[f, c].Value = "Si";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else
                        {

                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        c++;

                        if (complementosATR_PPH.TieneComplemento(r["CUPS22"].ToString().Substring(0, 20)))
                        {

                            total_passpool_horario++;
                            workSheet.Cells[f, c].Value = "Si";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else
                        {
                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        c++;

                        if (complementosATR_PPP.TieneComplemento(r["CUPS22"].ToString().Substring(0, 20)))
                        {

                            total_passpool_periodo++;
                            workSheet.Cells[f, c].Value = "Si";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else
                        {

                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        c++;

                        if (complementosATR_PS.TieneComplemento(r["CUPS22"].ToString().Substring(0, 20)))
                        {
                            total_passpool_subasta++;
                            workSheet.Cells[f, c].Value = "Si";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else
                        {
                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        c++;

                        if (r["tipo_contrato"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["tipo_contrato"].ToString();
                        }

                        c++;

                        if (r["gestionpropia_atr"] != System.DBNull.Value)
                        {
                            if (r["gestionpropia_atr"].ToString() == "Si")
                                total_gestion_atr++;
                            workSheet.Cells[f, c].Value = r["gestionpropia_atr"].ToString();
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        c++;

                        //if (r["TAM"] != System.DBNull.Value)
                        //{
                        //    workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //}else
                        //{
                        //    //valor_tam = tam.GetTamCups20(r["CUPS22"].ToString().Substring(0, 20));
                        //    valor_tam = tam.GetTAM(r["CUPS22"].ToString().Substring(0, 20));
                        //    if(valor_tam != -1111111)
                        //    {
                        //        workSheet.Cells[f, c].Value = valor_tam;
                        //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //    }
                        //    else
                        //    {
                        //        workSheet.Cells[f, c].Value = "Sin Facturas";                            
                        //    }

                        //}

                        // 230531 Comentado por cambio a TAM global
                        /*
                        valor_tam = tam.GetTAM(r["CUPS22"].ToString().Substring(0, 20));
                        if (valor_tam != -1111111)
                        {
                            workSheet.Cells[f, c].Value = valor_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        }
                        else
                        {
                            workSheet.Cells[f, c].Value = "Sin Facturas";
                        }

                        */
                        // 230531 ADDED por cambio a TAM global
                        if (r["TAM"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        }
                        else
                        {
                            workSheet.Cells[f, c].Value = "Sin Facturas";
                        }
                        // 230521 END ADDED
                        c++;

                        if (r["estado_contrato"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["estado_contrato"].ToString();
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            //if (!dic_agrup_estado_contrato.TryGetValue(r["estado_contrato"].ToString(), out o))
                            //    dic_agrup_estado_contrato.Add(r["estado_contrato"].ToString(), 1);
                            //else
                            //    o.total++;

                        }

                        c++;

                        if (r["Version"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["Version"]);

                        }

                        c++;

                        if (r["FPSERCON"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["FPSERCON"]);

                        }
                        c++;


                        if (r["revendedor"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["revendedor"].ToString();
                            if (r["revendedor"].ToString() == "Si")
                                total_revendedores++;

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        c++;


                        if (r["CONTREXT"] != System.DBNull.Value)
                        {

                            if (r["Version"] != System.DBNull.Value)
                            {
                                workSheet.Cells[f, c].Value
                               = Convertido(r["CONTREXT"].ToString(), Convert.ToInt32(r["Version"])) ? "Si" : "No";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            else
                                workSheet.Cells[f, c].Value = "No";



                        }

                        c++;

                        if (r["IDU"] != System.DBNull.Value)
                        {
                            if (perdidas.Medido_En_Baja(r["IDU"].ToString()))
                            {
                                total_medido_baja++;
                                workSheet.Cells[f, c].Value = "Si";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            else
                            {
                                workSheet.Cells[f, c].Value = "No";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                        }

                        c++;

                        if (r["IDU"] != System.DBNull.Value)
                        {
                            if (scea.Es_Multipunto(r["IDU"].ToString()))
                            {
                                total_multipuntos++;
                                workSheet.Cells[f, c].Value = "Si";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            else
                            {
                                workSheet.Cells[f, c].Value = "No";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                        }
                        c++;


                        string complemento = complementos_PS.GetCatalogo(r["IDU"].ToString());

                        switch (complemento)
                        {
                            case "Catálogo MT":
                                total_catalogo_mt++;
                                break;
                            case "Personalizado con revisión específica":
                                total_personalizado_r_esp++;
                                break;
                            case "Pesonalizado con revisión estándar":
                                total_personalizado_r_est++;
                                break;
                            case "Flexible":
                                total_flexible++;
                                break;
                        }

                        workSheet.Cells[f, c].Value = complemento;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        //Periodo pendiente
                        if (r["PERIODO_PENDIENTE"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["PERIODO_PENDIENTE"].ToString();
                        c++;
                        //Estado pendiente
                        if (r["cd_estado"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["cd_estado"].ToString();
                            if (r["de_estado"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = workSheet.Cells[f, c].Value + " " + r["de_estado"].ToString();
                        }
                        c++;
                        //Subestado pendiente
                        if (r["cd_subestado"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["cd_subestado"].ToString();
                            if (r["de_subestado"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = workSheet.Cells[f, c].Value + " " + r["de_subestado"].ToString();
                        }
                        c++;



                    }
                    db.CloseConnection();


                    #region guardado de totales
                    foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_empresa)
                        GuardaTotales(DateTime.Now, "EMPRESA", p.Key, p.Value.total);

                    foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_tarifa)
                        GuardaTotales(DateTime.Now, "TARIFA", p.Key, p.Value.total);

                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL AGRUPADAS", total_agrupada);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL AGE", total_age);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL AGORA", total_agora);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PASSTHOUGH", total_passthough);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PASSPOOL HORARIO", total_passpool_horario);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PASSPOOL SUBASTA", total_passpool_subasta);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL GESTIÓN PROPIA ATR", total_gestion_atr);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL REVENDEDORES", total_revendedores);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL MEDIDO EN BAJA", total_medido_baja);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL MULTIPUNTO", total_multipuntos);

                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL CATÁLOGO MT", total_catalogo_mt);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PERSONALIZADO CON REVISIÓN ESPECÍFICA", total_personalizado_r_esp);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PERSONALIZADO CON REVISIÓN ESTANDAR", total_personalizado_r_est);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL FLEXIBLE", total_flexible);



                    dic_totales = CargaTotales();

                    #endregion

                    var allCells = workSheet.Cells[1, 1, f, c];
                    headerCells = workSheet.Cells[1, 1, 1, c];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;
                    allCells = workSheet.Cells[1, 1, f, c];

                    allCells.AutoFitColumns();

                    workSheet.View.FreezePanes(2, 1);
                    workSheet.Cells["A1:AF1"].AutoFilter = true;
                    allCells.AutoFitColumns();


                    workSheet = excelPackage.Workbook.Worksheets.Add("AGRUPACIONES");
                    headerFont.Bold = true;
                    f = 1;
                    c = 1;

                    workSheet.Cells[f, c].Value = "EMPRESA";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    c++;
                    //workSheet.Cells[f, c].Value = "NUM";
                    //workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    //c++;                

                    foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_empresa)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.Key; c++;
                        c++;
                    }

                    f++;
                    f++;
                    c = 1;
                    workSheet.Cells[f, c].Value = "TARIFA";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    c++;


                    foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_tarifa)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.Key; c++;
                    }

                    f++;
                    f++;
                    c = 1;

                    #region Cabecera lateral Totales




                    workSheet.Cells[f, c].Value = "TOTAL AGRUPADAS:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    c++;


                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL AGE:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    c++;


                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL AGORA:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    c++;


                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL PASSTHOUGH:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    c++;


                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL PASSPOOL HORARIO:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    c++;
                    //workSheet.Cells[f, c].Value = total_passpool_horario;
                    //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL PASSPOOL PERIODO:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    c++;
                    //workSheet.Cells[f, c].Value = total_passpool_periodo;
                    //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    //GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PASSPOOL PERIODO", total_passpool_periodo);
                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL PASSPOOL SUBASTA:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    c++;
                    //workSheet.Cells[f, c].Value = total_passpool_subasta;
                    //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";                
                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL GESTIÓN PROPIA ATR:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    c++;
                    //workSheet.Cells[f, c].Value = total_gestion_atr;
                    //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";                
                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL REVENDEDORES:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    c++;
                    //workSheet.Cells[f, c].Value = total_revendedores;
                    //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";                
                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL MEDIDO EN BAJA:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    c++;
                    //workSheet.Cells[f, c].Value = total_medido_baja;
                    //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";                
                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL MULTIPUNTO:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL CATÁLOGO MT:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL CATÁLOGO PERS. REV. ESP.:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL CATÁLOGO PERS. REV. EST.:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);


                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = "TOTAL CATÁLOGO FLEXIBLE:";
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);


                    #endregion

                    firstOnly = true;
                    foreach (KeyValuePair<DateTime, List<EndesaEntity.facturacion.IRF_Totales_Hist>> p in dic_totales)
                    {
                        f = 1;

                        #region difenrecia
                        if (firstOnly)
                        {
                            c++;
                            firstOnly = false;
                            workSheet.Cells[f, c].Value = "DIF (D-C)";
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            f = 2;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "EMPRESA", workSheet.Cells[f, 1].Value.ToString()) -
                                GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f = 3;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "EMPRESA", workSheet.Cells[f, 1].Value.ToString()) -
                                GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f = 5;
                            foreach (KeyValuePair<string, EndesaEntity.Agrupacion> tarifa in dic_agrup_tarifa)
                            {

                                f++;
                                workSheet.Cells[f, c].Value =
                                    GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TARIFA", tarifa.Key) -
                                    GetTotal(p.Key, "TARIFA", tarifa.Key);
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            }

                            f = 16;

                            //f++;
                            //f++;

                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL AGRUPADAS") -
                                GetTotal(p.Key, "TOTALES", "TOTAL AGRUPADAS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL AGE") -
                                GetTotal(p.Key, "TOTALES", "TOTAL AGE");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL AGORA") -
                                GetTotal(p.Key, "TOTALES", "TOTAL AGORA");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL PASSTHOUGH") -
                                GetTotal(p.Key, "TOTALES", "TOTAL PASSTHOUGH");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL PASSPOOL HORARIO") -
                                GetTotal(p.Key, "TOTALES", "TOTAL PASSPOOL HORARIO");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL PASSPOOL PERIODO") -
                                GetTotal(p.Key, "TOTALES", "TOTAL PASSPOOL PERIODO");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL PASSPOOL SUBASTA") -
                                GetTotal(p.Key, "TOTALES", "TOTAL PASSPOOL SUBASTA");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL GESTIÓN PROPIA ATR") -
                                GetTotal(p.Key, "TOTALES", "TOTAL GESTIÓN PROPIA ATR");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL REVENDEDORES") -
                                GetTotal(p.Key, "TOTALES", "TOTAL REVENDEDORES");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL MEDIDO EN BAJA") -
                                GetTotal(p.Key, "TOTALES", "TOTAL MEDIDO EN BAJA");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL MULTIPUNTO") -
                                GetTotal(p.Key, "TOTALES", "TOTAL MULTIPUNTO");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL CATÁLOGO MT") -
                                GetTotal(p.Key, "TOTALES", "TOTAL CATÁLOGO MT");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL PERSONALIZADO CON REVISIÓN ESPECÍFICA") -
                                GetTotal(p.Key, "TOTALES", "TOTAL PERSONALIZADO CON REVISIÓN ESPECÍFICA");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";


                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL PERSONALIZADO CON REVISIÓN ESTANDAR") -
                                GetTotal(p.Key, "TOTALES", "TOTAL PERSONALIZADO CON REVISIÓN ESTANDAR");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL FLEXIBLE") -
                                GetTotal(p.Key, "TOTALES", "TOTAL FLEXIBLE");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";


                        }

                        #endregion

                        f = 1;

                        c++;
                        workSheet.Cells[f, c].Value = p.Key;
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                        f = 2;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f = 3;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f = 5;
                        foreach (KeyValuePair<string, EndesaEntity.Agrupacion> tarifa in dic_agrup_tarifa)
                        {

                            f++;
                            workSheet.Cells[f, c].Value = GetTotal(p.Key, "TARIFA", tarifa.Key);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        }

                        f = 16;

                        //f++;
                        //f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL AGRUPADAS");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL AGE");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL AGORA");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL PASSTHOUGH");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL PASSPOOL HORARIO");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL PASSPOOL PERIODO");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL PASSPOOL SUBASTA");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL GESTIÓN PROPIA ATR");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL REVENDEDORES");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL MEDIDO EN BAJA");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL MULTIPUNTO");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL CATÁLOGO MT");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL PERSONALIZADO CON REVISIÓN ESPECÍFICA");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL PERSONALIZADO CON REVISIÓN ESTANDAR");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL FLEXIBLE");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    }


                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();


                    excelPackage.Save();

                    if (param.GetValue("mail_tipologias") == "S")
                        EnvioCorreoInformeTipologias(ruta_salida_archivo);

                    
                    ss_pp.Update_Fecha_Fin("Facturación", "Informe Inventario Tipologias ES", "Informe Inventario Tipologias ES");

                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("InformeRevisionFacturas.Inventario_por_Tipologias: " + e.Message);

            }
        }


        private void GuardaTotales(DateTime fecha_informe, string grupo, string concepto, int cantidad)
        {
            MySQLDB db;
            MySqlCommand command;            
            string strSql;
            try
            {
                strSql = "replace into irf_totales_hist"
                    + " (fecha_informe, grupo, concepto, cantidad, last_update_date)"
                    + " values ('" + fecha_informe.ToString("yyyy-MM-dd") + "',"
                    + "'" + grupo + "'," 
                    + "'" + concepto + "'," 
                    + cantidad + ","
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
            
                Console.WriteLine(e.Message);
                ficheroLog.AddError("InformeRevisionFacturas.GuardaTotales: " + e.Message);
            }
        }

        private List<string> GetListaConceptos(string grupo)
        {
            Dictionary<string, string> d = new Dictionary<string, string>();
            foreach(KeyValuePair<DateTime, List<EndesaEntity.facturacion.IRF_Totales_Hist>> p in dic_totales)
            {
                for(int i = 0; i < p.Value.Count; i++)
                    if(p.Value[i].grupo == grupo)
                    {
                        string o;
                        if (!d.TryGetValue(p.Value[i].concepto, out o))
                            d.Add(p.Value[i].concepto, p.Value[i].concepto);
                    }

            }

            return d.Values.ToList();
        }

        private int GetTotal(DateTime fecha_informe, string grupo, string concepto)
        {
            int total = 0;
            List <EndesaEntity.facturacion.IRF_Totales_Hist> o;
            if (dic_totales.TryGetValue(fecha_informe, out o))
                foreach (EndesaEntity.facturacion.IRF_Totales_Hist p in o)
                    if (p.grupo == grupo && p.concepto == concepto)
                        total = p.cantidad;

            return total;
        }
        private Dictionary<DateTime, List<EndesaEntity.facturacion.IRF_Totales_Hist>> CargaTotales()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime fecha = new DateTime();

            Dictionary<DateTime, List<EndesaEntity.facturacion.IRF_Totales_Hist>> d =
                new Dictionary<DateTime, List<EndesaEntity.facturacion.IRF_Totales_Hist>>();

            try
            {

                strSql = "select fecha_informe, grupo, concepto, cantidad"
                    + " from irf_totales_hist"
                    + " where fecha_informe > '" + DateTime.Now.AddDays(-31).ToString("yyyy-MM-dd") + "'"
                    + " order by fecha_informe desc";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.IRF_Totales_Hist c = new EndesaEntity.facturacion.IRF_Totales_Hist();
                    fecha = Convert.ToDateTime(r["fecha_informe"]);
                    c.grupo = r["grupo"].ToString();
                    c.concepto = r["concepto"].ToString();
                    c.cantidad = Convert.ToInt32(r["cantidad"]);

                    List<EndesaEntity.facturacion.IRF_Totales_Hist> o;
                    if (!d.TryGetValue(fecha, out o))
                    {
                        o = new List<EndesaEntity.facturacion.IRF_Totales_Hist>();
                        o.Add(c);
                        d.Add(fecha, o);
                    }
                    else
                        o.Add(c);

                }
                db.CloseConnection();

                return d;
            }catch(Exception e)
            {
                return null;
            }
        }


        private bool Convertido(string contrato_externo, int version)
        {
            int o;
            if (dic_maxContrato.TryGetValue(contrato_externo, out o))
                return o < version;
            else
                return false;
        }

        public void Informe_Precios_PrePos_Circular()
        {

        }

        public void ImportarArchivo(string archivo, string tabla)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            string line = "";
            string[] campos;
            int i = 0;
            int total_registros = 0;
            int c = 0;
            string strSql = "";


            try
            {
                strSql = "delete from " + tabla;
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                System.IO.StreamReader file = new System.IO.StreamReader(archivo, System.Text.Encoding.GetEncoding(1252));
                while ((line = file.ReadLine()) != null)
                {
                    total_registros++;
                    campos = line.Split(';');
                    i++;
                    c = 1;

                    if (firstOnly)
                    {
                        sb.Append("replace into ").Append(tabla);
                        sb.Append(" (CCOUNIPS, CONTREXT, TESTCONT, FALTACON, FPSERCON, FPREVBAJ, FBAJACON,");
                        sb.Append(" CEMPTITU, CCONTRPS, CNUMSCPS, CLINNEG, CCLIENTE, CCALENPO, VDIAFACT, FSIGFACT,");
                        sb.Append(" FFINVESU, CTARIFA, CUPSREE, TPUNTMED, CCOMPOBJ, VNSEGHOR, VNUMTRAM, VPARAM01,");
                        sb.Append(" VPARAM02, VPARAM03, VPARAM04, VPARAM05, CCOMAUTO, TTICONPS, TPOTENCIP1, VPOTCALIP1,");
                        sb.Append(" TPOTENCIP2, VPOTCALIP2, TPOTENCIP3, VPOTCALIP3, TPOTENCIP4, VPOTCALIP4, TPOTENCIP5,");
                        sb.Append(" VPOTCALIP5, TPOTENCIP6, VPOTCALIP6) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(campos[c]).Append("',"); c++;
                    sb.Append("'").Append(campos[c]).Append("',"); c++;

                    for (int j = 1; j <= 6; j++)
                    {
                        sb.Append(CF(campos[c])).Append(","); c++;
                    }

                    sb.Append(CN(campos[c])).Append(","); c++; // CCONTRPS

                    for (int j = 1; j <= 7; j++)
                    {
                        sb.Append(CF(campos[c])).Append(","); c++;
                    }

                    sb.Append(CN(campos[c])).Append(","); c++; // CTARIFA
                    sb.Append(CN(campos[c])).Append(","); c++; // CUPSREE
                    sb.Append(CF(campos[c])).Append(","); c++; // TPUNTMED
                    sb.Append(CN(campos[c])).Append(","); c++; // CCOMPOBJ
                    sb.Append(CF(campos[c])).Append(","); c++; // VNSEGHOR
                    sb.Append(CF(campos[c])).Append(","); c++; // VNUMTRAM

                    for (int j = 1; j <= 5; j++)
                    {
                        sb.Append(CDouble(campos[c])).Append(","); c++;
                    }

                    sb.Append(CF(campos[c])).Append(","); c++; // CCOMAUTO
                    sb.Append(CN(campos[c])); c++; // TTICONPS

                    for (int j = 1; j <= 6; j++)
                    {
                        sb.Append(",").Append(CN(campos[c])); c++; // TPOTENCIP1
                        sb.Append(",").Append(CDouble(campos[c])); c++; // VPOTCALIP1
                    }

                    sb.Append("),");


                    if (i == 250)
                    {
                        Console.CursorLeft = 0;
                        Console.Write(total_registros);
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



            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }


        private bool ExisteEnAcumulada(List<EndesaEntity.facturacion.InformeRevisionFacturas> lista, 
            string cups22, double codSolAtr, string estadoSolAtr, string tipoSolAtr)
        {

            for(int i = 0; i < lista.Count; i++)
            {
                if(lista[i].cups22 == cups22 && 
                    lista[i].cod_sol_atr == codSolAtr && 
                    lista[i].estado_sol_atr == estadoSolAtr &&
                    lista[i].tipo_sol_atr == tipoSolAtr)
                    return true;
            }
            return false;
        }


        public void ImportaExtraccionAGE()
        {
            string line = "";
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            int numreg = 0;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            string[] c;
            int i = 0;
            string md5 = "";

            try
            {

                ficheroLog.Add("Ejecutando extractor: " + param.GetValue("Extractor_AGE", DateTime.Now, DateTime.Now));
                utilidades.Fichero.EjecutaComando(param.GetValue("Extractor_AGE", DateTime.Now, DateTime.Now));

                FileInfo archivo = new FileInfo(param.GetValue("Ubicacion_Informes")
                    + param.GetValue("archivo_extraccion_age"));


                md5 = utilidades.Fichero.checkMD5(archivo.FullName).ToString();

                if (archivo.Length > 0 && (md5 != param.GetValue("MD5_archivo_age")))
                {

                    strSql = "delete from irf_s20vge_tmp";
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    System.IO.StreamReader file = new System.IO.StreamReader(archivo.FullName, System.Text.Encoding.GetEncoding(1252));
                    while ((line = file.ReadLine()) != null)
                    {
                        numreg++;
                        c = line.Split(';');
                        if (firstOnly)
                        {
                            sb.Append("replace into irf_s20vge_tmp (CCOUNIPS, CEMPTITU, CONTREXT, CCONTRPS, version,");
                            sb.Append("TESTCONT, FPREALTA, FALTACON, FPSERCON, FPREVBAJ, FBAJACON, FINICON,");
                            sb.Append("FBAJACON2, FCARGA, TESTADO, CCONELEC) values ");
                            firstOnly = false;
                        }
                        i = 1;
                        sb.Append("('").Append(c[i]).Append("',"); i++; // CCOUNIPS
                        sb.Append(c[i]).Append(","); i++; // CEMPTITU
                        sb.Append("'").Append(c[i]).Append("',"); i++; // CONTREXT
                        sb.Append("'").Append(c[i]).Append("',"); i++; // CCONTRPS
                        sb.Append(c[i]).Append(","); i++; // version
                        sb.Append(c[i]).Append(","); i++; // TESTCONT

                        for(int j = 0; j < 8; j++)
                        { 
                            sb.Append(c[i]).Append(","); i++;                          
                        }

                        sb.Append(c[i]).Append(","); i++; // TESTADO
                        sb.Append("'").Append(c[i].Trim()).Append("'),"); i++; // CONTREXT



                        if (numreg == 250)
                        {
                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.CON);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            numreg = 0;
                        }
                    }
                    file.Close();

                    if (numreg > 0)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        numreg = 0;
                    }

                    strSql = "replace into irf_s20vge_hist select * from irf_s20vge";
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    strSql = "delete from irf_s20vge";
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();


                    strSql = "replace into irf_s20vge select * from irf_s20vge_tmp";
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    param.code = "MD5_archivo_age";
                    param.from_date = new DateTime(2021, 05, 01);
                    param.to_date = new DateTime(4999, 12, 31);
                    param.value = md5;
                    param.Save();

                }

            }
            catch(Exception e)
            {
                ficheroLog.AddError("ImportaExtraccionAGE --> " + e.Message);
            }

            

        }
             

        public void ImportaExtraccionAgrupadas()
        {
            string line = "";
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            int numreg = 0;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            string[] c;
            int i = 0;
            string md5_ES = "";
            string md5_PT = "";


            try
            {

                ficheroLog.Add("Ejecutando extractor: " + param.GetValue("Extractor_Agrupadas", DateTime.Now, DateTime.Now));
                utilidades.Fichero.EjecutaComando(param.GetValue("Extractor_Agrupadas", DateTime.Now, DateTime.Now));

                FileInfo archivo_ES = new FileInfo(param.GetValue("Ubicacion_Informes")
                    + param.GetValue("archivo_extraccion_agrupadas_ES"));

                FileInfo archivo_PT = new FileInfo(param.GetValue("Ubicacion_Informes")
                    + param.GetValue("archivo_extraccion_agrupadas_PT"));


                md5_ES = utilidades.Fichero.checkMD5(archivo_ES.FullName).ToString();
                md5_PT = utilidades.Fichero.checkMD5(archivo_PT.FullName).ToString();

                if ((archivo_ES.Length > 0 && (md5_ES != param.GetValue("MD5_archivo_agrupada_ES"))) ||
                    (archivo_PT.Length > 0 && (md5_PT != param.GetValue("MD5_archivo_agrupada_ES"))))
                {

                    strSql = "delete from irf_cont_fact_agrupada_tmp";
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    #region ES
                    System.IO.StreamReader file = new System.IO.StreamReader(archivo_ES.FullName, System.Text.Encoding.GetEncoding(1252));
                    while ((line = file.ReadLine()) != null)
                    {
                        numreg++;
                        c = line.Split(';');
                        if (firstOnly)
                        {
                            sb.Append("replace into irf_cont_fact_agrupada_tmp (cemptitu, ccomtcom, testcontcom, tfacagru, contraext,");
                            sb.Append("testconps, ctarifa, sergmerc, created_by, created_date) values ");
                            firstOnly = false;
                        }
                        i = 1;
                        sb.Append("(").Append(c[i]).Append(","); i++; // cemptitu
                        sb.Append("'").Append(c[i]).Append("',"); i++; // ccomtcom
                        sb.Append(c[i]).Append(","); i++; // testcontcom
                        sb.Append("'").Append(c[i]).Append("',"); i++; // tfacagru
                        sb.Append("'").Append(c[i]).Append("',"); i++; // contraext
                        sb.Append(c[i]).Append(","); i++; // testconps
                        sb.Append("'").Append(c[i]).Append("',"); i++; // ctarifa
                        sb.Append("'").Append(c[i].Trim()).Append("',"); i++; // sergmerc
                        sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                        sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");


                        if (numreg == 250)
                        {
                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.CON);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            numreg = 0;
                        }
                    }
                    file.Close();

                    if (numreg > 0)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        numreg = 0;
                    }
                    #endregion

                    #region PT

                    file = new System.IO.StreamReader(archivo_PT.FullName, System.Text.Encoding.GetEncoding(1252));
                    while ((line = file.ReadLine()) != null)
                    {
                        numreg++;
                        c = line.Split(';');
                        if (firstOnly)
                        {
                            sb.Append("replace into irf_cont_fact_agrupada_tmp (cemptitu, ccomtcom, testcontcom, tfacagru, contraext,");
                            sb.Append("testconps, ctarifa, sergmerc, created_by, created_date) values ");
                            firstOnly = false;
                        }
                        i = 1;
                        sb.Append("(").Append(c[i]).Append(","); i++; // cemptitu
                        sb.Append("'").Append(c[i]).Append("',"); i++; // ccomtcom
                        sb.Append(c[i]).Append(","); i++; // testcontcom
                        sb.Append("'").Append(c[i]).Append("',"); i++; // tfacagru
                        sb.Append("'").Append(c[i]).Append("',"); i++; // contraext
                        sb.Append(c[i]).Append(","); i++; // testconps
                        sb.Append("'").Append(c[i]).Append("',"); i++; // ctarifa
                        sb.Append("'").Append(c[i].Trim()).Append("',"); i++; // sergmerc
                        sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                        sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");



                        if (numreg == 250)
                        {
                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.CON);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            numreg = 0;
                        }
                    }
                    file.Close();

                    if (numreg > 0)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        numreg = 0;
                    }
                    #endregion





                    strSql = "replace into irf_cont_fact_agrupada_hist select * from irf_cont_fact_agrupada";
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    strSql = "delete from irf_cont_fact_agrupada";
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();


                    strSql = "replace into irf_cont_fact_agrupada select * from irf_cont_fact_agrupada_tmp";
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    param.code = "MD5_archivo_agrupada_ES";
                    param.from_date = new DateTime(2021, 05, 01);
                    param.to_date = new DateTime(4999, 12, 31);
                    param.value = md5_ES;
                    param.Save();

                    param.code = "MD5_archivo_agrupada_PT";
                    param.from_date = new DateTime(2021, 05, 01);
                    param.to_date = new DateTime(4999, 12, 31);
                    param.value = md5_PT;
                    param.Save();

                }

            }
            catch (Exception e)
            {
                ficheroLog.AddError("ImportaExtraccionAgrupadas --> " + e.Message);
            }



        }
               

        private void EnvioCorreoAgora(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);

            try
            {
                string from = param.GetValue("Buzon_envio_email");
                string to = param.GetValue("email_agora_para");
                string cc = param.GetValue("email_agora_copia");
                string subject = "Informe Revisión Factura Ágora: " + DateTime.Now.ToString("dd/MM/yyyy");
                string body = GeneraCuerpoHTML(CreaTabla(true), "&Aacute;gora");

                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");

                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (param.GetValue("EnviarMail") == "S")
                    //mes.SendMailWeb(from, to, cc, subject, body, archivo);
                    mes.SendMail(from, to, cc, subject, body, archivo);

                else
                    mes.SaveMail(from, to, cc, subject, body, archivo);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("Buzon_envio_email"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreoNoAgora: " + e.Message);
            }
        }


        private void EnvioCorreoNoAgora(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);

            try
            {
                string from = param.GetValue("Buzon_envio_email");
                string to = param.GetValue("email_noagora_para");
                string cc = param.GetValue("email_noagora_copia");
                string subject = "Informe Revisión Facturas No Ágora: " + DateTime.Now.ToString("dd/MM/yyyy");
                string body = GeneraCuerpoHTML(CreaTabla(false), "No &Aacute;gora");

                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (param.GetValue("EnviarMail") == "S")
                    mes.SendMail(from, to, cc, subject, body, archivo);                       
                    
                else
                    mes.SaveMail(from, to, cc, subject, body, archivo);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("Buzon_envio_email"));
            }catch(Exception e)
            {
                ficheroLog.AddError("EnvioCorreoNoAgora: " + e.Message);
            }
        }

        private void EnvioCorreoInformeTipologias(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                Console.WriteLine("Enviando correo Informe de Tipologias");

                string from = param.GetValue("Buzon_envio_email");
                string to = param.GetValue("email_tipologias_para");
                string cc = param.GetValue("email_tipologias_copia");
                string subject = param.GetValue("email_asunto_tipologias") + " " + DateTime.Now.ToString("dd/MM/yyyy");
                //string body = GeneraCuerpoHTML(CreaTabla(false), "No &Aacute;gora");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   Adjuntamos el informe de contratos por tipologías.");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("P.D:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   - La fecha de actualización de la extracción AGE es: "
                    + param.LastUpdateParameter("MD5_archivo_age").ToString("dd/MM/yyyy")).Append(".");                
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   - La fecha de actualización de la extracción Agrupadas es: "
                   + param.LastUpdateParameter("MD5_archivo_agrupada_ES").ToString("dd/MM/yyyy")).Append(".");

                // EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (param.GetValue("EnviarMail") == "S")
                    mes.SendMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);

                else
                    mes.SaveMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("Buzon_envio_email"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreoInformeTipologias: " + e.Message);
            }
        }

        private Dictionary<string, int> CargaUltimaVersion()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, int> d
                = new Dictionary<string, int>();

            DateTime fecha_anexion = new DateTime();            
            DateTime fecha_PSAT = new DateTime();


            try
            {
                Console.WriteLine("Analizando a Última versión de contratos PS");
                ficheroLog.Add("Analizando a Última versión de contratos PS");
                Console.WriteLine("");

                strSql = "select fecha from cont.ps_fechas_procesos pfp where pfp.proceso = 'PSAT'";
                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    fecha_PSAT = Convert.ToDateTime(r["fecha"]);
                    Console.WriteLine("ultima fecha tabla PS: " + fecha_PSAT.ToString("dd/MM/yyyy"));
                    Console.WriteLine("");
                }
                db.CloseConnection();

                if(fecha_PSAT.Date == DateTime.Now.Date)
                {
                    strSql = "select max(pah.Fecha_Anexion) fecha from cont.PS_AT_HIST pah"
                        + " where pah.Fecha_Anexion < '" + fecha_PSAT.ToString("yyyy-MM-dd") + "'";
                    ficheroLog.Add(strSql);
                    Console.WriteLine(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        fecha_anexion = Convert.ToDateTime(r["fecha"]);
                        Console.WriteLine("ultima fecha tabla PS_HIST: " + fecha_anexion.ToString("dd/MM/yyyy"));
                        Console.WriteLine("");
                    }
                    db.CloseConnection();
                }
                else
                {
                    strSql = "select max(pah.Fecha_Anexion) fecha from cont.PS_AT_HIST pah"
                         + " where pah.Fecha_Anexion < '" + fecha_PSAT.ToString("yyyy-MM-dd") + "'";
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        fecha_anexion = Convert.ToDateTime(r["fecha"]);
                        Console.WriteLine("ultima fecha tabla PS_HIST: " + fecha_anexion.ToString("dd/MM/yyyy"));
                        Console.WriteLine("");
                    }
                    db.CloseConnection();
                }


                Console.WriteLine("Utilizamos consulta de tabla para comparación");
                ficheroLog.Add("Utilizamos consulta de tabla para comparación");
                Console.WriteLine("");

                strSql = "select pah.CONTREXT, pah.Version"
                    + " from cont.PS_AT_HIST pah where"
                    + " pah.Fecha_Anexion = '" + fecha_anexion.ToString("yyyy-MM-dd") + "'"
                    + " and pah.CONTREXT is not null and pah.Version is not null"
                    + " group by pah.CONTREXT";

                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    d.Add(r["CONTREXT"].ToString(), Convert.ToInt32(r["Version"]));
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("CargaUltimaVersion: " + e.Message);
                return null;
            }
        }

        private List<EndesaEntity.facturacion.InformeRevisionFacturasTablaDias> CreaTabla(bool agora)
        {

            DateTime fd = new DateTime();
            DateTime fh = new DateTime();
            utilidades.Fechas f = new utilidades.Fechas();

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            int acumulados = 0;

            List<EndesaEntity.facturacion.InformeRevisionFacturasTablaDias> l 
                = new List<EndesaEntity.facturacion.InformeRevisionFacturasTablaDias>();

            try
            {
                fd = new DateTime(f.UltimoDiaHabil().Year, f.UltimoDiaHabil().Month, 1);
                

                strSql = "SELECT d.f_ult_mod, COUNT(*) as total FROM irf_acumulada d"
                    + " WHERE d.agora = '" + (agora ? "S" : "N") + "' and"
                    + " f_ult_mod >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " GROUP BY DATE(d.f_ult_mod)";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.InformeRevisionFacturasTablaDias c
                        = new EndesaEntity.facturacion.InformeRevisionFacturasTablaDias();

                    c.dia = Convert.ToDateTime(r["f_ult_mod"]).Date;
                    c.num_diarios = Convert.ToInt32(r["total"]);
                    acumulados += c.num_diarios;
                    c.num_acumulados = acumulados;
                    l.Add(c);
                }

                db.CloseConnection();

                strSql = "SELECT d.f_ult_mod, COUNT(*) as total FROM irf_diaria d"
                    + " WHERE d.agora = '" + (agora ? "S" : "N") + "'"                   
                    + " GROUP BY DATE(d.f_ult_mod)";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.InformeRevisionFacturasTablaDias c
                        = new EndesaEntity.facturacion.InformeRevisionFacturasTablaDias();

                    c.dia = Convert.ToDateTime(r["f_ult_mod"]).Date;
                    c.num_diarios = Convert.ToInt32(r["total"]);
                    acumulados += c.num_diarios;
                    c.num_acumulados = acumulados;
                    l.Add(c);
                }

                db.CloseConnection();

                return l;
            }catch(Exception e)
            {
                ficheroLog.AddError("CreaTabla: " + e.Message);
                return null;
            }
                        
        }
        private string GeneraCuerpoHTML(List<EndesaEntity.facturacion.InformeRevisionFacturasTablaDias> lista, string tipoInforme)
        {
            string body = "";
            string linea = "";

            try
            {
                body = param.GetValue("html_head");
                if (DateTime.Now.Hour > 14)
                    body = body.Replace("Buenos d&iacute;as.", "Buenas tardes:");
                body = body.Replace("||tipoinforme||", tipoInforme);
                for (int i = 0; i < lista.Count; i++)
                {
                    linea = param.GetValue("html_body");
                    linea = linea.Replace("||fecha||", lista[i].dia.ToString("dd/MM/yyyy"));
                    linea = linea.Replace("||diarias||", lista[i].num_diarios.ToString());
                    linea = linea.Replace("||acumuladas||", lista[i].num_acumulados.ToString());                    
                    body += linea;
                }

                body += param.GetValue("html_foot");

                return body;
            }
            catch (Exception e)
            {
                ficheroLog.AddError("GeneraCuerpoHTML: " + e.Message);
                return "";
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

