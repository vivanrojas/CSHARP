using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Style;
using Oracle.ManagedDataAccess.Client;
using Org.BouncyCastle.Asn1;
using Renci.SshNet.Common;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Telegram.Bot.Types;
using static Microsoft.Exchange.WebServices.Data.SearchFilter;

namespace EndesaBusiness.facturacion.puntos_calculo_btn
{
    public class Calculo_Prefacturas_BTN
    {

        logs.Log ficheroLog;
        utilidades.Param param;

      

        public Dictionary<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> dic { get; set; }
        Tarifas tarifa;
        Perfiles perfil;
        Ajustes ajustes;
        public Calculo_Prefacturas_BTN()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Calculo_Prefacturas_BTN");
            param = new utilidades.Param("lpc_btn_param", servidores.MySQLDB.Esquemas.FAC);
            dic = new Dictionary<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo>();
            tarifa = new Tarifas();
            perfil = new Perfiles();
            ajustes = new Ajustes();            
        }


        public void Proceso()
        {
            CargaDatos();
            //RellenaExcel();
        }

        private void CargaDatos()
        {
            ImportacionPlantillaExcel plantilla_excel = new ImportacionPlantillaExcel();
            Dictionary<string, int> dic_perfiles;
            Dictionary<string, int> dic_calendarios;

            string strSql = "";
            OracleServer ora_db;
            OracleCommand ora_command;
            OracleDataReader r;

            try
            {
                plantilla_excel.CargaDatos();
                //dic = plantilla_excel.dic_puntos_calculo;
                dic_perfiles = CargaPerfiles();
                dic_calendarios = CargaCalendarios();
                ajustes.Carga();


                strSql = "SELECT DISTINCT inv.TX_CPE, p.F_DESDE , p.F_HASTA, p.CONSUMO, p.F_GENERACION,"
                    + " NVL(e.TIPO,'PUNTO NORMAL') AS TIPO,"
                    + " INV.TX_TARIFA_ACCESO"
                    + " FROM MED_INF_BTN_SCE_PREFAC p"
                    + " INNER JOIN APL_INVENTARIO_PUNTOS_ACTIVOS inv ON"
                    + " inv.TX_CUPS_INT = p.CUPS"
                    + " LEFT JOIN MED_PUNTOS_ESPECIALES e ON e.CPE = inv.TX_CPE";
                ora_db = new OracleServer(OracleServer.Servidores.COMPOR);
                ora_command = new OracleCommand(strSql, ora_db.con);
                r = ora_command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo c =
                        new EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo();

                    //if (dic.TryGetValue(r["TX_CPE"].ToString(), out c))
                    //{
                    //    c.cpe = r["TX_CPE"].ToString();
                    //    c.f_desde = Convert.ToDateTime(r["F_DESDE"]);
                    //    c.f_hasta = Convert.ToDateTime(r["F_HASTA"]);
                    //    c.consumo = Convert.ToInt32(r["CONSUMO"]);
                    //    c.f_generacion = Convert.ToDateTime(r["F_GENERACION"]);
                    //    c.tarifa = tarifa.GetTipoTarifa(r["TX_TARIFA_ACCESO"].ToString());
                    //    c.tipo = r["TIPO"].ToString();
                    //    c.perfil = GetPerfil(dic_perfiles, c.cpe);
                    //    c.calendario = GetCalendario(dic_calendarios, c.cpe);
                    //    c.pct_aplicacion = ajustes.GetPCT_Ajuste(c.cpe);
                    //}

                    if (r["TX_CPE"] != System.DBNull.Value)
                        c.cpe = r["TX_CPE"].ToString();

                    if (r["F_DESDE"] != System.DBNull.Value)
                        c.f_desde = Convert.ToDateTime(r["F_DESDE"]);

                    if (r["F_HASTA"] != System.DBNull.Value)
                        c.f_hasta = Convert.ToDateTime(r["F_HASTA"]);

                    if (r["CONSUMO"] != System.DBNull.Value)
                        c.consumo = Convert.ToInt32(r["CONSUMO"]);

                    if (r["F_GENERACION"] != System.DBNull.Value)
                        c.f_generacion = Convert.ToDateTime(r["F_GENERACION"]);

                    if (r["F_GENERACION"] != System.DBNull.Value)
                        c.tarifa = tarifa.GetTipoTarifa(r["TX_TARIFA_ACCESO"].ToString());
                    if (r["F_GENERACION"] != System.DBNull.Value)
                        c.tipo = r["TIPO"].ToString();

                    if (r["F_GENERACION"] != System.DBNull.Value)
                        c.perfil = GetPerfil(dic_perfiles, c.cpe);

                    if (r["F_GENERACION"] != System.DBNull.Value)
                        c.calendario = GetCalendario(dic_calendarios, c.cpe);


                    c.pct_aplicacion = ajustes.GetPCT_Ajuste(c.cpe);


                    EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo o;
                    if (!dic.TryGetValue(c.cpe, out o))
                        dic.Add(c.cpe, c);

                }
                ora_db.CloseConnection();

                GuardaDatos_Prefacturas(dic.Values.ToList());

            }
            catch(Exception ex)
            {
                Console.WriteLine("Calculo_Prefactura_BTN.CargaDatos: " + ex.Message);
                ficheroLog.AddError("Calculo_Prefactura_BTN.CargaDatos: " + ex.Message);
            }

            
        }


        private void GuardaDatos_Prefacturas(List<EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> lista)
        {

            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            MySQLDB db;
            MySqlCommand command;
            int x = 0;
            string strSql = "";


            try
            {

                if (lista != null)
                {
                    strSql = "DELETE FROM lpc_btn_prefacturas";
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                }


                if (lista != null)
                    foreach (EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo p in lista)
                    {
                        x++;

                        if (firstOnly)
                        {
                            sb.Append("replace into lpc_btn_prefacturas (cpe, f_desde, f_hasta,");
                            sb.Append(" consumo, perfil, calendario, tarifa, aplicacion_pct,");                            
                            sb.Append(" created_by, created_date) values ");
                            firstOnly = false;
                        }

                        sb.Append("('").Append(p.cpe).Append("',");                        

                        if (p.f_desde > DateTime.MinValue)
                            sb.Append("'").Append(p.f_desde.ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (p.f_hasta > DateTime.MinValue)
                            sb.Append("'").Append(p.f_hasta.ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");
                        
                        sb.Append("'").Append(p.consumo.ToString().Replace(",",".")).Append("',");

                        sb.Append("'").Append(p.perfil).Append("',");
                        sb.Append("'").Append(p.calendario).Append("',");
                        sb.Append("'").Append(p.tarifa).Append("',");
                        sb.Append("'").Append(p.pct_aplicacion).Append("',");
                        sb.Append("'").Append(System.Environment.UserName).Append("',");
                        sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");


                        if (x == 500)
                        {
                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.FAC);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            x = 0;
                        }
                    }

                if (x > 0)
                {

                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    x = 0;
                }

                strSql = "REPLACE INTO lpc_btn_prefacturas_hist SELECT * FROM lpc_btn_prefacturas";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                    "Error en el guardado de datos",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private Dictionary<string, int> CargaPerfiles()
        {
            string strSql = "";
            OracleServer ora_db;
            OracleCommand ora_command;
            OracleDataReader r;
            string cpe = "";
            int perfil = 0;

            Dictionary<string, int> d = new Dictionary<string, int>();

            try
            {
                strSql = "select DISTINCT p.TX_CPE, R.R51100270, T.DESCRIPCION, R2.R51000040"
                     + " FROM APL_INVENTARIO_PUNTOS_ACTIVOS p"
                     + " INNER JOIN t_ges_mensajes T ON"
                     + " T.tx_cpe = p.TX_CPE"
                     + " INNER JOIN R51100000 R ON T.CD_ID = R.TX_ID And T.ID_SECUENCIAL = R.NM_SECUENCIAL"
                     + " INNER JOIN R51000000 R2 ON T.CD_ID = R2.TX_ID And"
                     + " T.ID_SECUENCIAL = R2.NM_SECUENCIAL"
                     + " left join T12520 T ON R.R51100270 = T.CODIGO"
                     + " WHERE R.R51100270 IS NOT NULL AND T.TX_PASO = 'P5100' order by R2.R51000040 DESC";
                ora_db = new OracleServer(OracleServer.Servidores.COMPOR);
                ora_command = new OracleCommand(strSql, ora_db.con);
                r = ora_command.ExecuteReader();
                while (r.Read())
                {
                    cpe = r["TX_CPE"].ToString();
                    perfil = Convert.ToInt32(r["R51100270"]);

                    int o;
                    if (!d.TryGetValue(cpe, out o))
                        d.Add(cpe, perfil);
                }
                ora_db.CloseConnection();               

                return d;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Calculo_Prefactura_BTN.CargaPerfiles: " + ex.Message);
                ficheroLog.AddError("Calculo_Prefactura_BTN.CargaPerfiles: " + ex.Message);
                return null;
            }
        }


        private string GetPerfil(Dictionary<string, int> dic,string cpe)
        {
           
            string codigo_perfil = "";
            int codigo = 0;

            try
            {
                if(dic.TryGetValue(cpe, out codigo))
                {
                    codigo_perfil = perfil.GetCodigoPerfil(codigo);
                }
                
                return codigo_perfil;

            }
            catch(Exception ex)
            {
                Console.WriteLine("Calculo_Prefactura_BTN.GetPerfil: " + ex.Message);
                ficheroLog.AddError("Calculo_Prefactura_BTN.GetPerfil: " + ex.Message);
                return null;
            }
        }


        private Dictionary<string, int> CargaCalendarios()
        {
            string strSql = "";
            OracleServer ora_db;
            OracleCommand ora_command;
            OracleDataReader r;
            string cpe = "";
            int calendario = 0;

            Dictionary<string, int> d = new Dictionary<string, int>();

            try
            {
                strSql = "select DISTINCT p.TX_CPE, R.R51120200, R2.R51000040" 
                     + " FROM APL_INVENTARIO_PUNTOS_ACTIVOS p"
                     + " INNER JOIN t_ges_mensajes T ON"
                     + " T.tx_cpe = p.TX_CPE"
                     + " INNER JOIN R51120000 R ON"
                     + " T.CD_ID = R.TX_ID And T.ID_SECUENCIAL = R.NM_SECUENCIAL"
                     + " INNER JOIN R51000000 R2 ON"
                     + " T.CD_ID = R2.TX_ID And T.ID_SECUENCIAL = R2.NM_SECUENCIAL"
                     + " WHERE R.R51120200 IS NOT NULL AND T.TX_PASO = 'P5100' order by R2.R51000040 DESC";
                ora_db = new OracleServer(OracleServer.Servidores.COMPOR);
                ora_command = new OracleCommand(strSql, ora_db.con);
                r = ora_command.ExecuteReader();
                while (r.Read())
                {
                    cpe = r["TX_CPE"].ToString();
                    calendario = Convert.ToInt32(r["R51120200"]);

                    int o;
                    if (!d.TryGetValue(cpe, out o))
                        d.Add(cpe, calendario);
                }
                ora_db.CloseConnection();

                return d;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Calculo_Prefactura_BTN.CargaCalendarios: " + ex.Message);
                ficheroLog.AddError("Calculo_Prefactura_BTN.CargaCalendarios: " + ex.Message);
                return null;
            }
        }

        private string GetCalendario(Dictionary<string, int> dic, string cpe)
        {            
            int codigo = 0;
            string calendario = "";

            try
            {
                if (dic.TryGetValue(cpe, out codigo))
                {
                    switch (codigo)
                    {
                        case 10:
                            calendario = "SIN CICLO";
                            break;
                        case 20:
                            calendario = "DIARIO";
                            break;
                        case 30:
                            calendario = "SEMANAL";
                            break;
                        case 40:
                            calendario = "SEMANAL";
                            break;
                        case 50:
                            calendario = "SEMANAL";
                            break;

                    }
                }

                return calendario;

            }
            catch(Exception ex)
            {
                ficheroLog.AddError("Calculo_Prefactura_BTN.GetCalendario: " + ex.Message);
                return null;
            }

        }


        public string RellenaExcel(string fichero)
        {

            int f = 1;
            int c = 1;

            try
            {

                //FileInfo filesave = new FileInfo(param.GetValue("ruta_salida_plantilla") + "\\" 
                //    + param.GetValue("ruta_salida_plantilla") + ".xlsx");

                if (fichero.Contains("xlsx"))
                    fichero = fichero.Replace("xlsx", "xlsm");

                FileInfo filesave = new FileInfo(fichero);

                // Ruta de la plantilla 
                FileInfo plantillaExcel = new FileInfo(System.Environment.CurrentDirectory + param.GetValue("plantilla_calculo_BTN"));

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);
                var workSheet = excelPackage.Workbook.Worksheets["PuntosCalculo"];

                List<EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> lista =
                   dic.Values.Where(z => z.f_desde > DateTime.MinValue).ToList();



                foreach(EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo p in lista)
                {
                    c = 1;
                    f++;
                    workSheet.Cells[f, c].Value = p.cpe; c++;
                    
                    workSheet.Cells[f, c].Value = p.f_desde;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    c++;

                    workSheet.Cells[f, c].Value = p.f_hasta;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    c++;

                    workSheet.Cells[f, c].Value = p.consumo;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    c++;

                    workSheet.Cells[f, c].Value = p.perfil;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; 
                    c++;


                    workSheet.Cells[f, c].Value = p.calendario;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;


                    workSheet.Cells[f, c].Value = p.tarifa;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    workSheet.Cells[f, c].Value = p.pct_aplicacion / 100;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0%";
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    c++;
                }

                excelPackage.SaveAs(filesave);
                excelPackage = null;

                return filesave.FullName;

            }
            catch(Exception ex)
            {
                ficheroLog.AddError("Calculo_Prefactura_BTN.RellenaExcel: " + ex.Message);
                return null;
            }
        }


        public void SalidaExcel(string fichero)
        {
            int f = 0;
            int c = 0;


            FileInfo fileInfo = new FileInfo(fichero);
            if (fileInfo.Exists)
                fileInfo.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Prefacturas BTN");

            var headerCells = workSheet.Cells[1, 1, 1, 6];
            var headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;
            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "CPE";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "F_DESDE";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "F_HASTA";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Consumo (kWh)";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "PERFIL";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "CALENDARIO";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "TARIFA";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "% Aplicación";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            foreach (KeyValuePair<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> p in dic)
            {
                f++;
                c = 1;


            }

            var allCells = workSheet.Cells[1, 1, f, c];
            workSheet.View.FreezePanes(2, 1);
            allCells.AutoFitColumns();
            excelPackage.Save();

        }


       
    }
}
