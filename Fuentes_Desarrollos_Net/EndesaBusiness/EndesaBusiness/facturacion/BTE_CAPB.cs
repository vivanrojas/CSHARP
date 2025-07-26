using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using System.Globalization;
using Oracle.ManagedDataAccess.Client;

namespace EndesaBusiness.facturacion
{
    public class BTE_CAPB
    {

        public List<EndesaEntity.facturacion.Fo_Tabla> lista { get; set; }
        public List<EndesaEntity.facturacion.Fo_Tabla> lista_prefacturas { get; set; }

        public BTE_CAPB()
        {
            lista = new List<EndesaEntity.facturacion.Fo_Tabla>();
            lista_prefacturas = new List<EndesaEntity.facturacion.Fo_Tabla>();
        }

        public void Facturas(string fd_emision, string fh_emision,
            string fd_consumo, string fh_consumo,
            string cups20)
        {

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            DateTime temp;


            try
            {

                lista = new List<EndesaEntity.facturacion.Fo_Tabla>();

                strSql = "DELETE FROM facturas_cap";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                strSql = "REPLACE INTO facturas_cap"
                    + " SELECT f.CUPSREE, f.CFACTURA, f.CREFEREN, f.SECFACTU, f.TESTFACT,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS, t.ICONFAC, f.VCUOVAFA, tf.descripcion AS TFACTURA"
                    + " FROM fo f"
                    + " INNER JOIN fo_tcon t ON"
                    + " t.CREFEREN = f.CREFEREN AND"
                    + " t.SECFACTU = f.SECFACTU AND"
                    + " t.TESTFACT = f.TESTFACT"
                    + " INNER JOIN fo_empresas e ON"
                    + " e.empresa_id = f.ID_ENTORNO"
                    + " LEFT OUTER JOIN fo_p_tipos_factura tf ON"
                    + " tf.id_tipo_factura = f.TFACTURA"
                    + " WHERE e.descripcion = 'BTE-Portugal'";

                if (cups20 != "")
                    strSql += " AND f.CUPSREE = '" + cups20 + "'";

                if (DateTime.TryParse(fd_emision, out temp) && DateTime.TryParse(fh_emision, out temp))
                    strSql += " AND (f.FFACTURA >= '" + Convert.ToDateTime(fd_emision).ToString("yyyy-MM-dd") + "' and"
                   + " f.FFACTURA <= '" + Convert.ToDateTime(fh_emision).ToString("yyyy-MM-dd") + "')";

                if (DateTime.TryParse(fd_consumo, out temp) && DateTime.TryParse(fh_consumo, out temp))
                    strSql += " AND (f.FFACTDES >= '" + Convert.ToDateTime(fd_consumo).ToString("yyyy-MM-dd") + "' and"
                   + " f.FFACTHAS <= '" + Convert.ToDateTime(fh_consumo).ToString("yyyy-MM-dd") + "')";
                
                strSql += " and t.TCONFAC = 1221"
                    + " AND f.TESTFACT IN ('N','Y','A')"
                    + " ORDER BY f.CUPSREE, f.FFACTDES, f.SECFACTU, f.TESTFACT desc";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "SELECT CUPS, CFACTURA, `Ref.` as ref, `Sec.` as sec,"
                    + "`Fecha emisión` as ffactura, `Periodo Desde` as ffactdes, `Periodo Hasta` as ffacthas, CAPB, consumo, TFACTURA"
                    + " FROM fact.facturas_cap WHERE TESTFACT <> 'A'";                
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.Fo_Tabla c = new EndesaEntity.facturacion.Fo_Tabla();
                    c.cupsree = r["CUPS"].ToString();
                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();

                    if (r["ref"] != System.DBNull.Value)
                        c.creferen = Convert.ToInt64(r["ref"]);

                    if (r["sec"] != System.DBNull.Value)
                        c.secfactu = Convert.ToInt32(r["sec"]);

                    if (r["ffactura"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["ffactura"]);

                    if (r["ffactdes"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["ffactdes"]);

                    if (r["ffacthas"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["ffacthas"]);

                    if (r["CAPB"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["CAPB"]);

                    if (r["consumo"] != System.DBNull.Value)
                        c.vcuovafa = Convert.ToInt32(r["consumo"]);

                    if (r["TFACTURA"] != System.DBNull.Value)
                        c.tfactura_descripcion = r["TFACTURA"].ToString();

                    lista.Add(c);

                }
                db.CloseConnection();                

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void Prefacturas()
        {
            OracleServer ora_db;
            OracleCommand ora_command;
            OracleDataReader r;
            string strSql = "";

            try 
            {

                strSql = "SELECT DISTINCT INV.TX_CPE,INV.TX_CONTRATO_ATR, P.F_DESDE, P.F_HASTA, P.CONSUMO, E.TIPO"
                    + " FROM MED_INF_BTE_SCE_INVENTARIO INV"
                    + " INNER JOIN MED_INF_BTE_SCE_PREFAC P ON INV.TX_CUPS_INT = P.CUPS"
                    + " LEFT JOIN MED_PUNTOS_ESPECIALES E ON INV.TX_CPE = E.CPE";
                ora_db = new OracleServer(OracleServer.Servidores.COMPOR);
                ora_command = new OracleCommand(strSql, ora_db.con);
                r = ora_command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.Fo_Tabla c = new EndesaEntity.facturacion.Fo_Tabla();
                    c.cupsree = r["TX_CPE"].ToString();

                    if (r["TX_CONTRATO_ATR"] != System.DBNull.Value)
                        c.crefaext = r["TX_CONTRATO_ATR"].ToString();

                    if (r["F_DESDE"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["F_DESDE"]);

                    if (r["F_HASTA"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["F_HASTA"]);

                    if (r["CONSUMO"] != System.DBNull.Value)
                        c.vcuovafa = Convert.ToInt32(r["CONSUMO"]);

                    c.comentarios = "10.A. Prefactura pendiente";

                    lista_prefacturas.Add(c);

                }
                ora_db.CloseConnection();

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void ExportExcel(string fichero)
        {

            int f = 0;
            int c = 0;


            FileInfo fileInfo = new FileInfo(fichero);
            if (fileInfo.Exists)
                fileInfo.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("BTE_CAPB");

            var headerCells = workSheet.Cells[1, 1, 1, 10];
            var headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;
            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "CUPS";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "FACTURA";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Ref.";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Sec.";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Consumo";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Fecha Emisión";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Periodo Desde";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Periodo Hasta";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "CAPB";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "TIPO FACTURA";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            foreach (EndesaEntity.facturacion.Fo_Tabla p in lista)
            {
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = p.cupsree;
                c++;

                workSheet.Cells[f, c].Value = p.cfactura;
                c++;

                workSheet.Cells[f, c].Value = p.creferen;
                c++;

                workSheet.Cells[f, c].Value = p.secfactu;
                c++;

                workSheet.Cells[f, c].Value = p.vcuovafa;
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                c++;

                if(p.ffactura > DateTime.MinValue)
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

                workSheet.Cells[f, c].Value = p.ifactura;
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                c++;

                if(p.tfactura_descripcion != "")
                {
                    workSheet.Cells[f, c].Value = p.tfactura_descripcion;                    
                }
                
                c++;

            }

            var allCells = workSheet.Cells[1, 1, f, 10];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:J1"].AutoFilter = true;
            allCells.AutoFitColumns();
            excelPackage.Save();

        }

        public void ExportExcelPrefacturas(string fichero)
        {

            int f = 0;
            int c = 0;


            FileInfo fileInfo = new FileInfo(fichero);
            if (fileInfo.Exists)
                fileInfo.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("BTE_CAPB");

            var headerCells = workSheet.Cells[1, 1, 1, 6];
            var headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;
            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "CUPS";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Contrato";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Fecha Desde";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Fecha Hasta";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Consumo";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;           

            workSheet.Cells[f, c].Value = "Estado";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            foreach (EndesaEntity.facturacion.Fo_Tabla p in lista_prefacturas)
            {
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = p.cupsree;
                c++;              

                workSheet.Cells[f, c].Value = p.crefaext;
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

                workSheet.Cells[f, c].Value = p.vcuovafa;
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                c++;

                workSheet.Cells[f, c].Value = p.comentarios;
                c++;
            }

            var allCells = workSheet.Cells[1, 1, f, c];
            workSheet.View.FreezePanes(2, 1);
            allCells.AutoFitColumns();
            excelPackage.Save();

        }

    }
}
