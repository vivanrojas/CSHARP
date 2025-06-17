using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.IO;

namespace EndesaBusiness.facturacion
{
    public class BTN_CAPB
    {
        public List<EndesaEntity.facturacion.Fo_Tabla> lista { get; set; }

        public BTN_CAPB()
        {
            lista = new List<EndesaEntity.facturacion.Fo_Tabla>();
        }

        public void Facturas(DateTime fd, DateTime fh, string cups20, bool agrupar)
        {

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            EndesaBusiness.compor.Inventario inventario =
                new compor.Inventario();

            try
            {
                if (agrupar)
                    strSql = "DELETE FROM facturas_cap_agrupadas";
                else
                    strSql = "DELETE FROM facturas_cap";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                if(agrupar)
                    strSql = "REPLACE INTO facturas_cap_agrupadas";
                else
                    strSql = "REPLACE INTO facturas_cap";

                strSql += " SELECT f.CUPSREE, f.CFACTURA, f.CREFEREN, f.SECFACTU,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS, t.ICONFAC, f.VCUOVAFA"
                    + " FROM fo f"
                    + " INNER JOIN fo_tcon t ON"
                    + " t.CREFEREN = f.CREFEREN AND"
                    + " t.SECFACTU = f.SECFACTU AND"
                    + " t.TESTFACT = f.TESTFACT"
                    + " INNER JOIN fo_empresas e ON"
                    + " e.empresa_id = f.ID_ENTORNO"
                    + " WHERE e.descripcion = 'BTN-Portugal' AND"
                    + " (f.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "') and";

                if (cups20 != "")
                    strSql += " f.CUPSREE = '" + cups20 + "' and";

                strSql += " t.TCONFAC = 1221" 
                    + " AND f.TESTFACT IN ('N','Y')"
                    + " order by f.CUPSREE, f.FFACTDES, f.SECFACTU";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "SELECT CUPS, CFACTURA, `Ref.` as ref, `Sec.` as sec,"
                    + "`Fecha emisión` as ffactura, `Periodo Desde` as ffactdes, `Periodo Hasta` as ffacthas, CAPB, consumo";
                if (agrupar)
                    strSql += " FROM fact.facturas_cap_agrupadas";
                else
                    strSql += " FROM fact.facturas_cap";

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

                    inventario.Get_Info_Inventario(c.cupsree);
                    if (inventario.existe)
                    {
                        c.perfil = inventario.perfil;
                        c.calendario = inventario.calendario;
                        c.tarifa = inventario.tarifa;
                    }

                    lista.Add(c);

                }
                db.CloseConnection();

            }
            catch (Exception ex)
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
            var workSheet = excelPackage.Workbook.Worksheets.Add("BTN_CAPB");

            var headerCells = workSheet.Cells[1, 1, 1, 12];
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

                workSheet.Cells[f, c].Value = p.ifactura;
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                c++;

                workSheet.Cells[f, c].Value = p.perfil; c++;
                workSheet.Cells[f, c].Value = p.calendario; c++;
                workSheet.Cells[f, c].Value = p.tarifa; c++;


            }

            var allCells = workSheet.Cells[1, 1, f, c];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:L1"].AutoFilter = true;
            allCells.AutoFitColumns();
            excelPackage.Save();

        }
    }
}
