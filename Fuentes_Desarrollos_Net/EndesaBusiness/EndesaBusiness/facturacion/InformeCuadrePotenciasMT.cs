using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
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
    public class InformeCuadrePotenciasMT
    {

        public List<EndesaEntity.facturacion.InformeCuadrePotenciasMT> listaFacturas =
            new List<EndesaEntity.facturacion.InformeCuadrePotenciasMT>();

        public InformeCuadrePotenciasMT()
        {

        }

        public InformeCuadrePotenciasMT(DateTime fd, DateTime fh, string cups20, string nif)
        {
            listaFacturas = BuscaFacturas(fd, fh, cups20, nif);
			

			
		}

        private List<EndesaEntity.facturacion.InformeCuadrePotenciasMT> BuscaFacturas(DateTime fd, DateTime fh, string cups20, string nif)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            List<EndesaEntity.facturacion.InformeCuadrePotenciasMT> lista
                = new List<EndesaEntity.facturacion.InformeCuadrePotenciasMT>();


            try
            {
                strSql = "SELECT f.DAPERSOC, f.CNIFDNIC, f.CFACTURA,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS, f.IFACTURA,"                    
                    + " 521 AS TCONFAC4,"
                    + " (SELECT tt.ICONFAC FROM fact.fo_tcon tt WHERE"
                    + " tt.CREFEREN = f.CREFEREN AND"
                    + " tt.SECFACTU = f.SECFACTU AND"
                    + " tt.TESTFACT = f.TESTFACT and"
                    + " tt.TCONFAC = 521) AS ICONFAC4,"
                    + " f.CUPSREE, f.POT_MAXIMA, NULL AS pc_correcta,"
                    + " DATEDIFF(f.FFACTHAS, f.FFACTDES) + 1 AS dias,"
                    + " NULL AS precio, NULL AS valor,"
                    + " (SELECT tt.ICONFAC FROM fact.fo_tcon tt WHERE"
                    + " tt.CREFEREN = f.CREFEREN AND"
                    + " tt.SECFACTU = f.SECFACTU AND"
                    + " tt.TESTFACT = f.TESTFACT and"
                    + " tt.TCONFAC = 521) AS facturado,"
                    + " NULL AS valor_a_devolver"
                    + " FROM fact.fo f INNER JOIN fact.fo_tcon t ON"
                    + " t.CREFEREN = f.CREFEREN AND"
                    + " t.SECFACTU = f.SECFACTU AND"
                    + " t.TESTFACT = f.TESTFACT"
                    + " WHERE (f.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "')";

                if (cups20 != "")
                    strSql += " and f.CUPSREE = '" + cups20 + "'";

                if (nif != "")
                    strSql += " and f.CNIFDNIC = '" + nif + "'";

                strSql += " AND f.TESTFACT IN ('N','Y')"
                    + " GROUP BY f.CREFEREN, f.SECFACTU, f.TESTFACT"
                    + " ORDER BY f.CREFEREN, f.FFACTDES, f.SECFACTU desc";


                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.InformeCuadrePotenciasMT c
                        = new EndesaEntity.facturacion.InformeCuadrePotenciasMT();

                    if (r["DAPERSOC"] != System.DBNull.Value)
                        c.dapersoc = r["DAPERSOC"].ToString();
                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        c.cnifdnic = r["CNIFDNIC"].ToString();
                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FFACTURA"]);
                    if (r["FFACTDES"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);
                    if (r["FFACTHAS"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);
                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"]);
                    
                    if (r["TCONFAC4"] != System.DBNull.Value)
                        c.tconfac4 = r["TCONFAC4"].ToString();
                    if (r["ICONFAC4"] != System.DBNull.Value)
                    {
                        c.iconfac4 = Convert.ToDouble(r["ICONFAC4"]);
                        c.facturado = c.iconfac4;
                    }
                        
                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cupsree = r["CUPSREE"].ToString();
                    if (r["POT_MAXIMA"] != System.DBNull.Value)
                        c.potencia_contratada = Convert.ToDouble(r["POT_MAXIMA"]);
                    if (r["pc_correcta"] != System.DBNull.Value)
                        c.potencia_correcta = Convert.ToDouble(r["pc_correcta"]);
                    if (r["dias"] != System.DBNull.Value)
                        c.dias = Convert.ToInt32(r["dias"]);

                    if(!lista.Exists(z => (z.cupsree == c.cupsree) && z.ffactdes == c.ffactdes))
                        lista.Add(c);

                }
                db.CloseConnection();

                return lista;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public void ExportarExcel(string fichero)
        {
            int f = 1;
            int c = 1;

            FileInfo fileInfo = new FileInfo(fichero);

            if (fileInfo.Exists)
                fileInfo.Delete();


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Facturas");

            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;
            

            workSheet.Cells[f, c].Value = "DAPERSOC"; c++;
            workSheet.Cells[f, c].Value = "CNIFDNIC"; c++;
            workSheet.Cells[f, c].Value = "CFACTURA"; c++;            
            workSheet.Cells[f, c].Value = "FFACTURA"; c++;
            workSheet.Cells[f, c].Value = "FFACTDES"; c++;
            workSheet.Cells[f, c].Value = "FFACTHAS"; c++;            
            workSheet.Cells[f, c].Value = "IFACTURA"; c++;            
            workSheet.Cells[f, c].Value = "TCONFAC4"; c++;
            workSheet.Cells[f, c].Value = "ICONFAC4"; c++;
            workSheet.Cells[f, c].Value = "CUPSREE"; c++;
            workSheet.Cells[f, c].Value = "POTENCIA_CONTRATADA"; c++;
            workSheet.Cells[f, c].Value = "PC Correcta"; c++;
            workSheet.Cells[f, c].Value = "dias"; c++;
            workSheet.Cells[f, c].Value = "preço"; c++;
            workSheet.Cells[f, c].Value = "Valor"; c++;
            workSheet.Cells[f, c].Value = "faturado"; c++;
            workSheet.Cells[f, c].Value = "Valor a devolver"; 

            foreach(EndesaEntity.facturacion.InformeCuadrePotenciasMT p in listaFacturas)
            {
                f++;
                c = 1;
                workSheet.Cells[f, c].Value = p.dapersoc; c++;
                workSheet.Cells[f, c].Value = p.cnifdnic; c++;
                workSheet.Cells[f, c].Value = p.cfactura; c++;
                workSheet.Cells[f, c].Value = p.ffactura; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.ffactdes; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.ffacthas; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.ifactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.##"; c++;                
                workSheet.Cells[f, c].Value = Convert.ToString(p.tconfac4); c++;
                workSheet.Cells[f, c].Value = p.iconfac4; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.##"; c++;
                workSheet.Cells[f, c].Value = p.cupsree; c++;
                workSheet.Cells[f, c].Value = p.potencia_contratada; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.##"; c++;
                c++; // Potencia correcta
                workSheet.Cells[f, c].Value = p.dias; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.iconfac4 / p.potencia_contratada / p.dias; // precio
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.0000"; c++;
                workSheet.Cells[f, c].Formula = "=L" + f + " * M" + f + " * N" + f; c++; // valor
                workSheet.Cells[f, c].Value = p.facturado; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.##"; c++;
                workSheet.Cells[f, c].Formula = "=P" + f + "-O" + f;

            }

            var allCells = workSheet.Cells[1, 1, f, c];
            var cellFont = allCells.Style.Font;
            cellFont.SetFromFont(new Font("Calibri", 8));
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:Q1"].AutoFilter = true;
            allCells.AutoFitColumns();


            excelPackage.Save();

        }


    }
}
