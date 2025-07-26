using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class ConsumosPortugal
    {
        Dictionary<string, EndesaEntity.facturacion.InformeFacturasConsumoPortugal> dic;
        List<EndesaEntity.facturacion.InformeFacturasConsumoPortugal> lista_facturas;
        public ConsumosPortugal(DateTime fecha_consumo_desde, DateTime fecha_consumo_hasta, bool sacar_facturas, bool sacar_descuadre)
        {
            dic = new Dictionary<string, EndesaEntity.facturacion.InformeFacturasConsumoPortugal>();
            lista_facturas = new List<EndesaEntity.facturacion.InformeFacturasConsumoPortugal>();
            Carga(fecha_consumo_desde, fecha_consumo_hasta);
            InformeExcel(fecha_consumo_desde, fecha_consumo_hasta, sacar_facturas, sacar_descuadre);

        }
        
        private void Carga(DateTime ffactdes, DateTime ffacthas)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            try
            {
                Console.WriteLine("Consultando Facturas ...");

                strSql = "SELECT e.descripcion AS Empresa," 
                    + " f.CNIFDNIC, f.DAPERSOC, f.CUPSREE,"
                    + " f.CREFEREN, f.SECFACTU, f.TESTFACT,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " f.VCUOVAFA, f.CFACTURA,"
                    + " f.VCONSACP AS PUNTA,"
                    + " f.VCONSACL AS LLANO,"
                    + " f.VCONSACV AS VALLE,"
                    + " f.SUPERVALLE,"
                    + " f.VCONATH1 AS P1,"
                    + " f.VCONATH2 AS P2,"
                    + " f.VCONATH3 AS P3,"
                    + " f.VCONATH4 AS P4,"
                    + " f.VCONATH5 AS P5,"
                    + " f.VCONATH6 AS P6"
                    + " from fo f INNER JOIN fo_empresas e ON"
                    + " e.empresa_id = f.ID_ENTORNO"
                    + " WHERE"                    
                    + " (f.FFACTDES <= '" + ffacthas.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTHAS >= '" + ffactdes.ToString("yyyy-MM-dd")   + "')"
                    + " and e.descripcion IN ('BTN-Portugal','BTE-Portugal','MT-Portugal')"
                    + " and f.TIPONEGOCIO = 'L'";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.InformeFacturasConsumoPortugal c = new EndesaEntity.facturacion.InformeFacturasConsumoPortugal();

                    if (r["Empresa"] != System.DBNull.Value)
                        c.empresa = r["Empresa"].ToString();

                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        c.nif = r["CNIFDNIC"].ToString();

                    if (r["DAPERSOC"] != System.DBNull.Value)
                        c.razon_social = r["DAPERSOC"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cupsree = r["CUPSREE"].ToString();

                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();

                    if (r["CREFEREN"] != System.DBNull.Value)
                        c.creferen = Convert.ToInt64(r["CREFEREN"]);

                    if (r["SECFACTU"] != System.DBNull.Value)
                        c.secfactu = Convert.ToInt32(r["SECFACTU"]);

                    if (r["TESTFACT"] != System.DBNull.Value)
                        c.testfact = r["TESTFACT"].ToString();

                    if (r["FFACTDES"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                    if (r["FFACTHAS"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);


                    if (r["FFACTDES"] != System.DBNull.Value)
                        c.min_ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                    if (r["FFACTHAS"] != System.DBNull.Value)
                        c.max_ffacthas = Convert.ToDateTime(r["FFACTHAS"]);

                    if (r["VCUOVAFA"] != System.DBNull.Value)
                        c.vcuovafa = Convert.ToInt32(r["VCUOVAFA"]);
                   

                    if(c.empresa == "BTN-Portugal" || c.empresa == "BTE-Portugal")
                    {
                        if(c.vcuovafa < 0)
                        {
                            if (r["PUNTA"] != System.DBNull.Value)                            
                                c.p1 = Math.Abs(Convert.ToInt32(r["PUNTA"])) * -1;                                
                            if (r["LLANO"] != System.DBNull.Value)
                                c.p2 = Math.Abs(Convert.ToInt32(r["LLANO"])) * -1;
                            if (r["VALLE"] != System.DBNull.Value)
                                c.p3 = Math.Abs(Convert.ToInt32(r["VALLE"])) * -1;
                            if (r["SUPERVALLE"] != System.DBNull.Value)
                                c.p4 = Math.Abs(Convert.ToInt32(r["SUPERVALLE"])) * -1;
                        }
                        else
                        {
                            if(r["PUNTA"] != System.DBNull.Value)
                                c.p1 = Math.Abs(Convert.ToInt32(r["PUNTA"]));
                            if (r["LLANO"] != System.DBNull.Value)
                                c.p2 = Math.Abs(Convert.ToInt32(r["LLANO"]));
                            if (r["VALLE"] != System.DBNull.Value)
                                c.p3 = Math.Abs(Convert.ToInt32(r["VALLE"]));
                            if (r["SUPERVALLE"] != System.DBNull.Value)
                                c.p4 = Math.Abs(Convert.ToInt32(r["SUPERVALLE"]));
                        }

                    }
                    else
                    {
                        if (c.vcuovafa < 0)
                        {
                            if (r["P1"] != System.DBNull.Value)
                                c.p1 = Math.Abs(Convert.ToInt32(r["P1"])) * -1;
                            if (r["P2"] != System.DBNull.Value)
                                c.p2 = Math.Abs(Convert.ToInt32(r["P2"])) * -1;
                            if (r["P3"] != System.DBNull.Value)
                                c.p3 = Math.Abs(Convert.ToInt32(r["P3"])) * -1;
                            if (r["P4"] != System.DBNull.Value)
                                c.p4 = Math.Abs(Convert.ToInt32(r["P4"])) * -1;
                            if (r["P5"] != System.DBNull.Value)
                                c.p5 = Math.Abs(Convert.ToInt32(r["P5"])) * -1;
                            if (r["P6"] != System.DBNull.Value)
                                c.p6 = Math.Abs(Convert.ToInt32(r["P6"])) * -1;
                        }
                        else
                        {
                            if (r["P1"] != System.DBNull.Value)
                                c.p1 = Math.Abs(Convert.ToInt32(r["P1"]));
                            if (r["P2"] != System.DBNull.Value)
                                c.p2 = Math.Abs(Convert.ToInt32(r["P2"]));
                            if (r["P3"] != System.DBNull.Value)
                                c.p3 = Math.Abs(Convert.ToInt32(r["P3"]));
                            if (r["P4"] != System.DBNull.Value)
                                c.p4 = Math.Abs(Convert.ToInt32(r["P4"]));
                            if (r["P5"] != System.DBNull.Value)
                                c.p5 = Math.Abs(Convert.ToInt32(r["P5"]));
                            if (r["P6"] != System.DBNull.Value)
                                c.p6 = Math.Abs(Convert.ToInt32(r["P6"]));
                        }
                    }

                    lista_facturas.Add(c);

                    EndesaEntity.facturacion.InformeFacturasConsumoPortugal o;
                    if (!dic.TryGetValue(c.empresa + c.cupsree, out o))                                            
                        dic.Add(c.empresa + c.cupsree, c);                                         
                    else
                    {
                        o.vcuovafa += c.vcuovafa;
                        o.p1 += c.p1;
                        o.p2 += c.p2;
                        o.p3 += c.p3;
                        o.p4 += c.p4;
                        o.min_ffactdes = c.min_ffactdes < o.min_ffactdes ? c.min_ffactdes : o.min_ffactdes;
                        o.max_ffacthas = c.max_ffacthas > o.max_ffacthas ? c.max_ffacthas : o.max_ffacthas;
                    }
                    
                }
                db.CloseConnection();
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        private void InformeExcel(DateTime ffactdes, DateTime ffacthas, bool sacar_facturas, bool sacar_descuadre)
        {
            int f = 1;
            int c = 1;
            int i_factura = 0;
            

            string fichero = @"c:\Temp\MT_BTE_BTN_consumos_desde_" 
                +  ffactdes.ToString("yyyyMMdd") + "_al_"
                + ffacthas.ToString("yyyyMMdd") 
                + ffactdes.ToString("yyyyMMdd") + "_"

                +  DateTime.Now.ToString("yyyyMMdd_HHmmss") 
                + ".xlsx";


            FileInfo fileInfo = new FileInfo(fichero);

            if (fileInfo.Exists)
                fileInfo.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("CONSUMOS");

            var headerCells = workSheet.Cells[1, 1, 1, 38];
            var headerFont = headerCells.Style.Font;
            f = 1;

            headerFont.Bold = true;

            #region Cabecera_Excel

            workSheet.Cells[f, c].Value = "Empresa";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "CIF";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Razón Social";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "CPE";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "VCUOVAFA";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "P1";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "P2";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "P3";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "P4";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "P5";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "P6";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;
            


            workSheet.Cells[f, c].Value = "Fecha mínima";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Fecha máxima";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;



            #endregion


            foreach (KeyValuePair<string, EndesaEntity.facturacion.InformeFacturasConsumoPortugal> p in dic)
            {
                c = 1;
                f++;
                workSheet.Cells[f, c].Value = p.Value.empresa; c++;
                workSheet.Cells[f, c].Value = p.Value.nif; c++;
                workSheet.Cells[f, c].Value = p.Value.razon_social; c++;
                workSheet.Cells[f, c].Value = p.Value.cupsree; c++;
                workSheet.Cells[f, c].Value = p.Value.vcuovafa; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.p1; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.p2; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.p3; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.p4; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.p5; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.p6; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                workSheet.Cells[f, c].Value = p.Value.min_ffactdes;
                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                workSheet.Cells[f, c].Value = p.Value.max_ffacthas;
                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
            }
                        

            var allCells = workSheet.Cells[1, 1, f, c];
            var cellFont = allCells.Style.Font;
            cellFont.SetFromFont(new Font("Calibri", 8));
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:S1"].AutoFilter = true;
            allCells.AutoFitColumns();


            if (sacar_facturas)
            {

                workSheet = excelPackage.Workbook.Worksheets.Add("FACTURAS");
                headerCells = workSheet.Cells[1, 1, 1, 38];
                headerFont = headerCells.Style.Font;
                f = 1;
                c = 1;

                workSheet.Cells[f, c].Value = "Empresa";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CREFEREN";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "SECFACTU";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "TESTFACT";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CPE";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CFACTURA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "FECHA FACTURA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "FECHA DESDE";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "FECHA HASTA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "VCUOVAFA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "P1";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "P2";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "P3";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "P4";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "P5";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "P6";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;


                Console.WriteLine();

                

                    foreach (EndesaEntity.facturacion.InformeFacturasConsumoPortugal p in lista_facturas)
                    {
                        c = 1;
                        f++;
                        i_factura++;

                        Console.CursorLeft = 0;
                        Console.Write("Volcando factura: " + i_factura + " /  " + lista_facturas.Count());

                        workSheet.Cells[f, c].Value = p.empresa; c++;
                        workSheet.Cells[f, c].Value = p.creferen; c++;
                        workSheet.Cells[f, c].Value = p.secfactu; c++;
                        workSheet.Cells[f, c].Value = p.testfact; c++;
                        workSheet.Cells[f, c].Value = p.cupsree; c++;
                        workSheet.Cells[f, c].Value = p.cfactura; c++;

                        if (p.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;


                        if (p.ffactdes > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.ffactdes;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;


                        if (p.ffacthas > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.ffacthas;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.vcuovafa; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                        workSheet.Cells[f, c].Value = p.p1; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                        workSheet.Cells[f, c].Value = p.p2; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                        workSheet.Cells[f, c].Value = p.p3; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                        workSheet.Cells[f, c].Value = p.p4; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                        workSheet.Cells[f, c].Value = p.p5; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                        workSheet.Cells[f, c].Value = p.p6; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    }


                allCells = workSheet.Cells[1, 1, f, c];
                cellFont = allCells.Style.Font;
                cellFont.SetFromFont(new Font("Calibri", 8));
                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:P1"].AutoFilter = true;
                allCells.AutoFitColumns();
            }

            if (sacar_descuadre)
            {

                workSheet = excelPackage.Workbook.Worksheets.Add("FACTURAS_DESCUADRES");
                headerCells = workSheet.Cells[1, 1, 1, 38];
                headerFont = headerCells.Style.Font;
                f = 1;
                c = 1;

                #region cabecera_descuadres

                workSheet.Cells[f, c].Value = "Empresa";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CREFEREN";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "SECFACTU";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "TESTFACT";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CPE";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CFACTURA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "FECHA FACTURA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "FECHA DESDE";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "FECHA HASTA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "VCUOVAFA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "P1";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "P2";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "P3";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "P4";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "P5";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "P6";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                #endregion

                Console.WriteLine();

                foreach (KeyValuePair<string, EndesaEntity.facturacion.InformeFacturasConsumoPortugal> d in dic)
                {
                    if(d.Value.vcuovafa != (d.Value.p1 + d.Value.p2 + d.Value.p3 + d.Value.p4 + d.Value.p5 + d.Value.p6))
                    {
                        List<EndesaEntity.facturacion.InformeFacturasConsumoPortugal> sub_lista =
                            lista_facturas.Where(z => z.empresa == d.Value.empresa && z.cupsree == d.Value.cupsree).ToList();

                        foreach (EndesaEntity.facturacion.InformeFacturasConsumoPortugal p in sub_lista)
                        {
                            c = 1;
                            f++;
                            i_factura++;

                            Console.CursorLeft = 0;
                            Console.Write("Volcando factura: " + i_factura + " /  " + lista_facturas.Count());

                            workSheet.Cells[f, c].Value = p.empresa; c++;
                            workSheet.Cells[f, c].Value = p.creferen; c++;
                            workSheet.Cells[f, c].Value = p.secfactu; c++;
                            workSheet.Cells[f, c].Value = p.testfact; c++;
                            workSheet.Cells[f, c].Value = p.cupsree; c++;
                            workSheet.Cells[f, c].Value = p.cfactura; c++;

                            if (p.ffactura > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.ffactura;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;


                            if (p.ffactdes > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.ffactdes;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;


                            if (p.ffacthas > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.ffacthas;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;

                            workSheet.Cells[f, c].Value = p.vcuovafa; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                            workSheet.Cells[f, c].Value = p.p1; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                            workSheet.Cells[f, c].Value = p.p2; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                            workSheet.Cells[f, c].Value = p.p3; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                            workSheet.Cells[f, c].Value = p.p4; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                            workSheet.Cells[f, c].Value = p.p5; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                            workSheet.Cells[f, c].Value = p.p6; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                        }
                    }
                }

                    


                allCells = workSheet.Cells[1, 1, f, c];
                cellFont = allCells.Style.Font;
                cellFont.SetFromFont(new Font("Calibri", 8));
                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:P1"].AutoFilter = true;
                allCells.AutoFitColumns();
            }

            excelPackage.Save();


        }

    }
}
