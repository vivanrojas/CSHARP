using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using EndesaBusiness.utilidades;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf.codec.wmf;
using Microsoft.Graph.CallRecords;
using OfficeOpenXml.Style;
using EndesaEntity.sigame;

namespace EndesaBusiness.facturacion
{
    public class ConsumptionDataExtration
    {

        logs.Log ficheroLog;
        List<EndesaEntity.facturacion.ConsumptionDataExtration> lista_ES;
        List<EndesaEntity.facturacion.ConsumptionDataExtration> lista_PT;

        public ConsumptionDataExtration()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_ConsumptionDataExtration");
        }

        public ConsumptionDataExtration(DateTime fd, DateTime fh)
        {
            lista_ES = new List<EndesaEntity.facturacion.ConsumptionDataExtration>();
            lista_PT = new List<EndesaEntity.facturacion.ConsumptionDataExtration>();
            CargaDatos(fd, fh);
        }


        private void CargaDatos(DateTime fd, DateTime fh)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;


            strSql = "SELECT anio, mes, mercado, linea, segmento, consumo FROM consumption_data_report"
                + " WHERE pais = 'ES' AND"
                + " (anio >= " + fd.Year + " and mes >= " + fd.Month + ") and"
                + " (anio <= " + fh.Year + " and mes <= " + fh.Month + ")"
                + " order by anio, mes, linea, segmento";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.facturacion.ConsumptionDataExtration c =
                    new EndesaEntity.facturacion.ConsumptionDataExtration();

                c.anio = Convert.ToInt32(r["anio"]);
                c.mes = Convert.ToInt32(r["mes"]);
                c.mercado = r["mercado"].ToString();
                c.linea = r["linea"].ToString();
                c.segmento = r["segmento"].ToString();
                c.consumo = Convert.ToInt32(r["consumo"]);

                lista_ES.Add(c);


            }
            db.CloseConnection();

            strSql = "SELECT anio, mes, mercado, linea, segmento, consumo FROM consumption_data_report"
                + " WHERE pais = 'PT' AND"
                + " (anio >= " + fd.Year + " and mes >= " + fd.Month + ") and"
                + " (anio <= " + fh.Year + " and mes <= " + fh.Month + ")"
                + " order by anio, mes, linea, segmento";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.facturacion.ConsumptionDataExtration c =
                    new EndesaEntity.facturacion.ConsumptionDataExtration();

                c.anio = Convert.ToInt32(r["anio"]);
                c.mes = Convert.ToInt32(r["mes"]);
                c.mercado = r["mercado"].ToString();
                c.linea = r["linea"].ToString();
                c.segmento = r["segmento"].ToString();
                c.consumo = Convert.ToInt32(r["consumo"]);

                lista_PT.Add(c);


            }
            db.CloseConnection();
        }



        public void GeneraExcel(string rutaFichero)
        {
            int f = 1;
            int c = 1;

            
            int mes = 0;
            FileInfo fileInfo;

            bool firstOnly = true;

            fileInfo = new FileInfo(rutaFichero);
            if (fileInfo.Exists)
                fileInfo.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Portugal");


            var headerCells = workSheet.Cells[1, 1, 1, 8];
            var headerFont = headerCells.Style.Font;

            var allCells = workSheet.Cells[1, 1, 1, 8];
            var cellFont = allCells.Style.Font;
            cellFont.Bold = true;

            workSheet.Cells[f, c].Value = "Año"; 
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Value = "Mes";
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Value = "Mercado";
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Value = "Línea"; 
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Value = "Segmento"; 
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Value = "Consumo (MWh)";
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;


            foreach (EndesaEntity.facturacion.ConsumptionDataExtration pc in lista_PT)
            {
                

                if (firstOnly)
                {                    
                    mes = pc.mes;
                    firstOnly = false;
                }

                if(mes != pc.mes)
                {
                    c = 1;
                    workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
                    workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
                    workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
                    workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
                    workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
                    workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
                    firstOnly = true;
                }


                f++;
                c = 1;

                workSheet.Cells[f, c].Value = pc.anio; c++;
                workSheet.Cells[f, c].Value = pc.mes; c++;
                workSheet.Cells[f, c].Value = pc.mercado; c++;
                workSheet.Cells[f, c].Value = pc.linea; c++;
                workSheet.Cells[f, c].Value = pc.segmento; c++;
                workSheet.Cells[f, c].Value = pc.consumo;
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;



            }

            c = 1;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;

            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:F1"].AutoFilter = true;
            allCells = workSheet.Cells[1, 1, f, c];
            allCells.AutoFitColumns();


            workSheet = excelPackage.Workbook.Worksheets.Add("España");

            f = 1;
            c = 1;

            headerCells = workSheet.Cells[1, 1, 1, 8];
            headerFont = headerCells.Style.Font;

            allCells = workSheet.Cells[1, 1, 1, 8];
            cellFont = allCells.Style.Font;
            cellFont.Bold = true;

            workSheet.Cells[f, c].Value = "Año";
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Value = "Mes";
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Value = "Mercado";
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Value = "Línea";
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Value = "Segmento";
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Value = "Consumo (MWh)";
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;


            firstOnly = true;

            foreach (EndesaEntity.facturacion.ConsumptionDataExtration pc in lista_ES)
            {

                if (firstOnly)
                {
                    mes = pc.mes;
                    firstOnly = false;
                }

                if (mes != pc.mes)
                {
                    c = 1;
                    workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
                    workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
                    workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
                    workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
                    workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
                    workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
                    firstOnly = true;
                }

                f++;
                c = 1;

                workSheet.Cells[f, c].Value = pc.anio; c++;
                workSheet.Cells[f, c].Value = pc.mes; c++;
                workSheet.Cells[f, c].Value = pc.mercado; c++;
                workSheet.Cells[f, c].Value = pc.linea; c++;
                workSheet.Cells[f, c].Value = pc.segmento; c++;
                workSheet.Cells[f, c].Value = pc.consumo;
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
            }

            c = 1;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; c++;

            //var cellFont = allCells.Style.Font;
            //cellFont.SetFromFont(new Font("Calibri", 8));
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:F1"].AutoFilter = true;
            allCells = workSheet.Cells[1, 1, f, c];
            allCells.AutoFitColumns();
            excelPackage.Save();
            
        }


        public void ConstruyeDatos()
        {
            DateTime fd = new DateTime();
            DateTime fh = new DateTime();
            MySQLDB db;
            MySqlCommand command;            
            string strSql = "";

            // Siemple calculamos sobre mes actual - 1

            fh = DateTime.Now;
            fh = new DateTime(fh.Year, fh.Month, 1);
            fd = fh.AddMonths(-1);

            Console.WriteLine("Consumption Data Extraction");
            Console.WriteLine("===========================");
            Console.WriteLine("");


            strSql = "DELETE FROM consumption_data_extraction";
            ficheroLog.Add(strSql);
            Console.WriteLine(strSql);
            Console.WriteLine("");
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


            strSql = "REPLACE INTO consumption_data_extraction"
                + " SELECT e.descripcion AS EMPRESA, f.CNIFDNIC, f.DAPERSOC, f.FFACTURA,"
                + " f.FFACTDES, f.FFACTHAS,"
                + " 'FM' AS Mercado, f.TIPONEGOCIO,"
                + " if (substr(f.CNIFDNIC, 1, 1) = 'P' or"
                + " substr(f.CNIFDNIC, 1, 1) = 'Q' or"             
                + " substr(f.CNIFDNIC, 1, 1) = 'S', 'B2G', 'B2B') AS Segmento,"
                + " f.VCUOVAFA AS Consumo, f.TFACTURA, f.TESTFACT,"
                + "'" + System.Environment.UserName + "', NOW(), NULL, NOW()"
                + " FROM fact.fo f"
                + " INNER JOIN fact.fo_empresas e ON"
                + " e.empresa_id = f.ID_ENTORNO"
                + " WHERE"
                + " e.descripcion IN ('MT-España','EEXXI')"
                + " AND f.FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' AND"
                + " f.FFACTURA < '" + fh.ToString("yyyy-MM-dd") + "'";
            ficheroLog.Add(strSql);
            Console.WriteLine(strSql);
            Console.WriteLine("");
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "REPLACE INTO consumption_data_extraction"
                + " SELECT e.descripcion AS EMPRESA, f.CNIFDNIC, f.DAPERSOC, f.FFACTURA," 
                + " f.FFACTDES, f.FFACTHAS,"
                + " 'FM' AS Mercado, f.TIPONEGOCIO, if (aapp.CNIFDNIC IS NULL, 'B2B','B2G') AS Segmento,"
                + " f.VCUOVAFA AS Consumo, f.TFACTURA, f.TESTFACT,"
                + "'" + System.Environment.UserName + "', NOW(), NULL, NOW()"
                + " FROM fact.fo f"
                + " INNER JOIN fact.fo_empresas e ON"
                + " e.empresa_id = f.ID_ENTORNO"
                + " LEFT OUTER JOIN fact.fo_aapp_portugal aapp ON"
                + " aapp.CNIFDNIC = f.CNIFDNIC"
                + " WHERE"
                + " e.descripcion IN('BTN-Portugal', 'BTE-Portugal','MT-Portugal')"
                + " AND f.FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' AND"
                + " f.FFACTURA < '" + fh.ToString("yyyy-MM-dd") + "'";
            ficheroLog.Add(strSql);
            Console.WriteLine(strSql);
            Console.WriteLine("");
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


            strSql = "REPLACE INTO consumption_data_report"
                + " SELECT 'ES' AS PAIS, YEAR(f.FFACTURA) AS YEAR, MONTH(f.FFACTURA) AS MONTH,"
                + " f.Mercado, 'GAS ESPAÑA' AS Línea, f.Segmento,"
                + " round(SUM(f.Consumo) / 1000, 0) AS Consumo,"
                + "'" + System.Environment.UserName + "', NOW(), NULL, NOW()"
                + " FROM fact.consumption_data_extraction f"
                + " WHERE f.TFACTURA NOT IN(5, 6) AND f.TIPONEGOCIO = 'G'"
                + " AND f.EMPRESA IN('MT-España','EEXXI')"
                + " GROUP BY YEAR(f.FFACTURA), MONTH(f.FFACTURA), f.Segmento";
            ficheroLog.Add(strSql);
            Console.WriteLine(strSql);
            Console.WriteLine("");
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "REPLACE INTO consumption_data_report"
                + " SELECT 'ES' AS PAIS, YEAR(f.FFACTURA) AS YEAR, MONTH(f.FFACTURA) AS MONTH,"
                + " f.Mercado, 'ELECTRICIDAD ESPAÑA' AS Línea, f.Segmento,"
                + " round(SUM(f.Consumo) / 1000, 0) AS Consumo," 
                + "'" + System.Environment.UserName + "', NOW(), NULL, NOW()"
                + " FROM fact.consumption_data_extraction f"
                + " WHERE f.TFACTURA NOT IN(5, 6) AND f.TIPONEGOCIO = 'L'"
                + " AND f.EMPRESA IN('MT-España','EEXXI')"
                + " GROUP BY YEAR(f.FFACTURA), MONTH(f.FFACTURA), f.Segmento";
            ficheroLog.Add(strSql);
            Console.WriteLine(strSql);
            Console.WriteLine("");
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "REPLACE INTO consumption_data_report"
                + " SELECT 'PT' AS PAIS, YEAR(f.FFACTURA) AS YEAR, MONTH(f.FFACTURA) AS MONTH,"
                + " f.Mercado, 'ELECTRICIDAD PORTUGAL' AS Línea, f.Segmento,"
                + " round(SUM(f.Consumo) / 1000, 0) AS Consumo,"
                + "'" + System.Environment.UserName + "', NOW(), NULL, NOW()"
                + " FROM fact.consumption_data_extraction f"
                + " WHERE f.TFACTURA NOT IN(5, 6) AND f.TIPONEGOCIO = 'L'"
                + " AND f.EMPRESA IN('BTN-Portugal','BTE-Portugal','MT-Portugal')"
                + " GROUP BY YEAR(f.FFACTURA), MONTH(f.FFACTURA), f.Segmento";
            ficheroLog.Add(strSql);
            Console.WriteLine(strSql);
            Console.WriteLine("");
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = " REPLACE INTO consumption_data_report"
                + " SELECT 'PT' AS PAIS, YEAR(f.FFACTURA) AS YEAR, MONTH(f.FFACTURA) AS MONTH,"
                + " f.Mercado, 'GAS PORTUGAL' AS Línea, f.Segmento,"
                + " round(SUM(f.Consumo) / 1000, 0) AS Consumo,"
                + "'" + System.Environment.UserName + "', NOW(), NULL, NOW()"
                + " FROM fact.consumption_data_extraction f"
                + " WHERE f.TFACTURA NOT IN(5, 6) AND f.TIPONEGOCIO = 'G'"
                + " AND f.EMPRESA IN('BTN-Portugal','BTE-Portugal','MT-Portugal')"
                + " GROUP BY YEAR(f.FFACTURA), MONTH(f.FFACTURA), f.Segmento";
            ficheroLog.Add(strSql);
            Console.WriteLine(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

    }
}

