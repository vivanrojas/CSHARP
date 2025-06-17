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
using System.Threading;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class Pendiente_Mes_Gas_Factura
    {
        // Informe diario sobre una consulta del pendiente de facturar de Gas

        List<EndesaEntity.facturacion.Pendiente_Mes_Gas_Factura> lista;
        utilidades.Param param;
        utilidades.Seguimiento_Procesos ss_pp;
        public Pendiente_Mes_Gas_Factura()
        {
            param = new utilidades.Param("ppg_param", MySQLDB.Esquemas.FAC);
            ss_pp = new utilidades.Seguimiento_Procesos();
            
        }

        public void GenerarInforme()
        {
            ss_pp.Update_Fecha_Inicio("Facturación", "Informe PDTE Facturación GAS", "Informe PDTE Facturación GAS");
            lista = Carga();
            InformeExcel(lista);
            ss_pp.Update_Fecha_Fin("Facturación", "Informe PDTE Facturación GAS", "Informe PDTE Facturación GAS");
        }


        private List<EndesaEntity.facturacion.Pendiente_Mes_Gas_Factura> Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            List<EndesaEntity.facturacion.Pendiente_Mes_Gas_Factura> l =
                new List<EndesaEntity.facturacion.Pendiente_Mes_Gas_Factura>();

            try
            {
                strSql = "SELECT c.CUPS, c.IdPtoSuministro, c.CIF," 
                    + " i.DAPERSOC AS cliente, c.Tarifa AS ATR,"
                    + " c.TOP, mf.Mes,"
                    + " mf.Medida, mf.Comentario," 
                    + " mf.Facturacion, c.Distribuidora, mf.Carga"                    
                    + " FROM med.GestGas_Clientes c"
                    + " INNER JOIN med.GestGas_medidasfacturas mf ON"
                    + " c.IdPtoSuministro = mf.`IdPto Medida`"
                    + " INNER JOIN fact.cm_inventario_gas i ON"
                    + " i.ID_PS = mf.`IdPto Medida`"
                    + " WHERE(mf.Mes IS NOT NULL"
                    + " AND mf.Mes <> '200605' AND mf.Mes <> '200604')"
                    + " AND mf.Medida IS NOT NULL"
                    + " AND mf.Facturacion IS NULL"
                    + " AND (mf.Comentario IS NULL OR mf.Comentario = '')"
                    + " GROUP BY mf.`IdPto Medida`, mf.Mes"
                    + " ORDER BY  mf.Medida; ";

                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.Pendiente_Mes_Gas_Factura c = 
                        new EndesaEntity.facturacion.Pendiente_Mes_Gas_Factura();

                    c.cups = r["CUPS"].ToString();
                    c.id_ps = Convert.ToInt32(r["IdPtoSuministro"]);
                    c.cif = r["CIF"].ToString();
                    c.cliente = r["cliente"].ToString();
                    c.atr = r["ATR"].ToString();
                    c.top = r["TOP"].ToString() == "SI";
                    c.mes = Convert.ToInt32(r["Mes"]);
                    if (r["Medida"] != System.DBNull.Value)
                        c.medida = Convert.ToDateTime(r["Medida"]);
                    if (r["Comentario"] != System.DBNull.Value)
                        c.comentario = r["Comentario"].ToString();
                    if (r["Facturacion"] != System.DBNull.Value)
                        c.facturacion = r["Facturacion"].ToString();
                    if (r["Distribuidora"] != System.DBNull.Value)
                        c.distribuidora = r["Distribuidora"].ToString();
                    if (r["Carga"] != System.DBNull.Value)
                        c.carga = Convert.ToDateTime(r["Carga"]);

                    l.Add(c);



                }
                db.CloseConnection();
                return l;
            }
            catch(Exception e)
            {
                return null;
            }
            
        }


        private void InformeExcel(List<EndesaEntity.facturacion.Pendiente_Mes_Gas_Factura> lista)
        {
            int f = 1;
            int c = 1;

            string fichero = param.GetValue("salida_informe",DateTime.Now, DateTime.Now)
                + param.GetValue("nombre_informe", DateTime.Now, DateTime.Now)
                + DateTime.Now.ToString("yyyyMMdd_HHmmss")
                + ".xlsx";

            FileInfo fileInfo = new FileInfo(fichero);

            if (fileInfo.Exists)
                fileInfo.Delete();

            string[] listaArchivos = 
                Directory.GetFiles(param.GetValue("salida_informe"), param.GetValue("nombre_informe") + "*.*");
            for (int j = 0; j < listaArchivos.Length; j++)
            {

                fileInfo = new FileInfo(listaArchivos[j]);
                fileInfo.Delete();
            }


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Lista Facturación");

            var headerCells = workSheet.Cells[1, 1, 1, 12];
            var headerFont = headerCells.Style.Font;
            f = 1;

            headerFont.Bold = true;

            #region Cabecera Excel
            workSheet.Cells[f, c].Value = "CUPS";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "IdPtoSuministro";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "CIF";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "CLIENTE";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "ATR";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "TOP";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Mes";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Medida";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Comentario";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Facturación";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Distribuidora";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Carga";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;


            #endregion


            foreach(EndesaEntity.facturacion.Pendiente_Mes_Gas_Factura p in lista)
            {
                c = 1;
                f++;

                workSheet.Cells[f, c].Value = p.cups; c++;
                workSheet.Cells[f, c].Value = p.id_ps; 
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";  c++;
                workSheet.Cells[f, c].Value = p.cif; c++;
                workSheet.Cells[f, c].Value = p.cliente; c++;
                workSheet.Cells[f, c].Value = p.atr; c++;
                workSheet.Cells[f, c].Value = (p.top ? "SI" : "NO"); c++;
                workSheet.Cells[f, c].Value = p.mes; c++;
                workSheet.Cells[f, c].Value = p.medida;
                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.comentario; c++;
                workSheet.Cells[f, c].Value = p.facturacion; c++;
                workSheet.Cells[f, c].Value = p.distribuidora; c++;
                workSheet.Cells[f, c].Value = p.carga;
                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                
            }

            var allCells = workSheet.Cells[1, 1, f, c];
            var cellFont = allCells.Style.Font;
            cellFont.SetFromFont(new Font("Calibri", 8));
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:L1"].AutoFilter = true;
            allCells.AutoFitColumns();
            excelPackage.Save();
            excelPackage = null;
            GeneraMail(fileInfo.FullName);
            Thread.Sleep(5000);
            

        }

        private void GeneraMail(string adjunto)
        {
            string body = "";
            string subject = "";
            string from = "";
            string to = "";
            string cc = null;            

            body = (DateTime.Now.Hour > 14 ? "Buenas tardes:" : "Buenos días:")
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "     Se adjunta informe."
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "Un saludo.";

            from = param.GetValue("mail_from", DateTime.Now, DateTime.Now);
            to = param.GetValue("mail_to", DateTime.Now, DateTime.Now);
            subject = param.GetValue("mail_subject", DateTime.Now, DateTime.Now);
            cc = param.GetValue("remitente", DateTime.Now, DateTime.Now);

            //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
            EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
            mes.SendMail(from, to, cc, subject, body, adjunto);
        }
    }
}
