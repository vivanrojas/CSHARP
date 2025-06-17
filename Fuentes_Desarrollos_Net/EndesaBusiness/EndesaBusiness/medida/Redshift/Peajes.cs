using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using Microsoft.Graph;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data.Odbc;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.medida.Redshift
{
    public class Peajes
    {
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;
        logs.Log ficheroLog;
        public Peajes()
        {
            ss_pp = new utilidades.Seguimiento_Procesos();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_Copia_Facturas_Peajes_BI");
        }


        private DateTime UltimaFechaCopiado()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime fecha = new DateTime();

            strSql = "SELECT max(fh_act_dmco) AS fh_act_dmco FROM t_ed_h_facts_atr_ml";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
                fecha = Convert.ToDateTime(r["fh_act_dmco"]);
            db.CloseConnection();
            Console.WriteLine("Última fecha de copiado RedShift: " + fecha.ToString("dd/MM/yyyy"));
            ficheroLog.Add("Última fecha de copiado RedShift: " + fecha.ToString("dd/MM/yyyy"));

            return fecha;
        }




        public void Copia_Peajes()
        {

            MySQLDB dbmy;
            MySqlCommand commandmy;
            MySqlDataReader rmy;
            string strSql = "";
            int i = 0;
            int total_cups = 0;

            DateTime fd = new DateTime();
            DateTime fh = new DateTime();

            Console.WriteLine("Inicio del Proceso");
            Console.WriteLine("==================");

            ficheroLog.Add("Inicio del Proceso");
            ficheroLog.Add("==================");

            fd = UltimaFechaCopiado();

            Dictionary<string, string> dic_cups =
                    new Dictionary<string, string>();

            List<string> lista_cups13 = new List<string>();
            bool firstOnly = false;

            ss_pp.Update_Fecha_Inicio("Medida", "Copia Peajes BI", "Copia Peajes BI");

            // Buscamos facturas peajes del inventario anterior a una semana
            strSql = "SELECT CUPS20 FROM med_listado_scea_cups_historico h"
                + " WHERE FECHAANEXION < '" + DateTime.Now.AddDays(-7).ToString("yyyy-MM-dd") + "'"
                + " GROUP BY CUPS20";
            dbmy = new MySQLDB(MySQLDB.Esquemas.MED);
            commandmy = new MySqlCommand(strSql, dbmy.con);
            rmy = commandmy.ExecuteReader();
            while (rmy.Read())
            {

                if (firstOnly)
                {
                    if (dic_cups.Count > 0)
                        dic_cups.Clear();

                    firstOnly = false;
                }

                i++;
                total_cups++;
                if (rmy["CUPS20"] != System.DBNull.Value)
                {
                    string o;
                    if(!dic_cups.TryGetValue(rmy["CUPS20"].ToString(), out o))
                    {
                        dic_cups.Add(rmy["CUPS20"].ToString(), rmy["CUPS20"].ToString());
                    }
                    
                    lista_cups13.Add(rmy["CUPS20"].ToString());
                }
                                
                if(i == 500)
                {
                    i = 0;
                    RecorreQuery_cups20(fd.Date, dic_cups);
                    firstOnly = true;
                }

            }
            dbmy.CloseConnection();
            //Console.WriteLine("Encontrados med_listado_scea_cups_historico: " + String.Format("{0:N0}", dic_cups.Count));

            if(i > 0)
            {
                RecorreQuery_cups20(fd.Date, dic_cups);
            }


            // Buscamos facturas peajes del inventario nuevo
            firstOnly = true;
            fd = fd.AddYears(-1);
            strSql = "SELECT CUPS20 FROM med_listado_scea_cups_historico h"
                + " WHERE FECHAANEXION >= '" + DateTime.Now.AddDays(-7).ToString("yyyy-MM-dd") + "'"
                + " GROUP BY CUPS20";
            dbmy = new MySQLDB(MySQLDB.Esquemas.MED);
            commandmy = new MySqlCommand(strSql, dbmy.con);
            rmy = commandmy.ExecuteReader();
            while (rmy.Read())
            {

                if (firstOnly)
                {
                    if (dic_cups.Count > 0)
                        dic_cups.Clear();

                    firstOnly = false;
                }

                i++;
                total_cups++;
                if (rmy["CUPS20"] != System.DBNull.Value)
                {
                    string o;
                    if (!dic_cups.TryGetValue(rmy["CUPS20"].ToString(), out o))
                    {
                        dic_cups.Add(rmy["CUPS20"].ToString(), rmy["CUPS20"].ToString());
                    }

                    lista_cups13.Add(rmy["CUPS20"].ToString());
                }

                if (i == 500)
                {
                    i = 0;
                    RecorreQuery_cups20(fd.Date, dic_cups);
                    firstOnly = true;
                }

            }
            dbmy.CloseConnection();
            //Console.WriteLine("Encontrados med_listado_scea_cups_historico: " + String.Format("{0:N0}", dic_cups.Count));

            if (i > 0)
            {
                RecorreQuery_cups20(fd.Date, dic_cups);
            }


            ss_pp.Update_Fecha_Fin("Medida", "Copia Peajes BI", "Copia Peajes BI");

        }

        public void Copia_Peajes_Fechas(DateTime fd, DateTime fh)
        {
            StringBuilder sb = new StringBuilder();
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;

            MySQLDB dbmy;
            MySqlCommand commandmy;
            MySqlDataReader rmy;
            string strSql = "";

            bool firstOnly = true;
            int j = 0;
            int k = 0;
            int totalRegistros = 0;
            string o;

            try
            {

                Dictionary<string, string> dic_cups =
                    new Dictionary<string, string>();

                strSql = "SELECT CUPS13 FROM med_listado_scea_cups_historico h"
                + " GROUP BY CUPS13";
                dbmy = new MySQLDB(MySQLDB.Esquemas.MED);
                commandmy = new MySqlCommand(strSql, dbmy.con);
                rmy = commandmy.ExecuteReader();
                while (rmy.Read())
                {
                    if (rmy["CUPS13"] != System.DBNull.Value)
                        dic_cups.Add(rmy["CUPS13"].ToString(), rmy["CUPS13"].ToString());
                }
                dbmy.CloseConnection();
                Console.WriteLine("Encontrados med_listado_scea_cups_historico: " + String.Format("{0:N0}", dic_cups.Count));




                //foreach (KeyValuePair<string, string> p in dic_cups)
                //{


                    //Console.WriteLine("Contanto registros BI para " + p.Key);
                    //db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                    //command = new OdbcCommand(Consulta_Count(fd, p.Key), db.con);
                    //r = command.ExecuteReader();
                    //while (r.Read())
                    //{
                    //    totalRegistros = Convert.ToInt32(r["total_registos"]);
                    //}
                    //db.CloseConnection();

                    //Console.WriteLine("Total registros BI: " + String.Format("{0:N0}", totalRegistros));


                    if (totalRegistros > 0)
                    {

                        
                    }
                

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }




        public string ExportExcel(string fichero, List<EndesaEntity.medida.PuntoSuministro> lista_cups, bool fecha_consumo)
        {
            int f = 0;
            int c = 0;


            


            try
            {

                List<EndesaEntity.medida.Peajes_Vista> lista_facturas = Consulta_Peajes_MySQL(lista_cups, fecha_consumo);


                FileInfo fileInfo = new FileInfo(fichero);
                if (fileInfo.Exists)
                    fileInfo.Delete();

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("Facturas");

                var headerCells = workSheet.Cells[1, 1, 1, 95];
                var headerFont = headerCells.Style.Font;
                f = 1;
                c = 1;
                headerFont.Bold = true;

                #region Cabecera

                workSheet.Cells[f, c].Value = "CUPS13";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "CUPS20";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Cod_Factura";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Estado_Factura";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Tipo_Factura";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_Factura";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_Desde";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_Hasta";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Ind_Cuadre_Periodo_c";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "FFM_AE_total_c";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "FFM_R1_total_c";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "FFM_PMax_total_c";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "FFF_AE_total_c";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "FFF_R1_total_c";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "FFF_PFac_total_c";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "FFF_Imp_Exc_Pot_total_c";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "FFF_Imp_Exc_Pot";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "FFF_Imp_Exc_R1";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                for (int i = 1; i <= 6; i++)
                {
                    workSheet.Cells[f, c].Value = "FFM_AE_" + i;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    PintaRecuadro(excelPackage, f, c); c++;
                }

                for (int i = 1; i <= 6; i++)
                {
                    workSheet.Cells[f, c].Value = "FFM_R1_" + i;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    PintaRecuadro(excelPackage, f, c); c++;
                }

                for (int i = 1; i <= 6; i++)
                {
                    workSheet.Cells[f, c].Value = "FFM_PMax_" + i;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    PintaRecuadro(excelPackage, f, c); c++;
                }

                for (int i = 1; i <= 6; i++)
                {
                    workSheet.Cells[f, c].Value = "FFF_AE_" + i;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    PintaRecuadro(excelPackage, f, c); c++;
                }

                for (int i = 1; i <= 6; i++)
                {
                    workSheet.Cells[f, c].Value = "FFF_R1_" + i;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    PintaRecuadro(excelPackage, f, c); c++;
                }


                workSheet.Cells[f, c].Value = "FFF_R4_6";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                for (int i = 1; i <= 6; i++)
                {
                    workSheet.Cells[f, c].Value = "FFF_PFac" + i;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    PintaRecuadro(excelPackage, f, c); c++;
                }

                for (int i = 1; i <= 6; i++)
                {
                    workSheet.Cells[f, c].Value = "FFF_Imp_Exc_Pot_" + i;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    PintaRecuadro(excelPackage, f, c); c++;
                }

                workSheet.Cells[f, c].Value = "CUPS20_Metra";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Procedencia";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Ind_Perdidas";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Porcentaje_Perdidas";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Pot_Trafo_Perdidas_VA";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Tipo_PM";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Ind_Autoconsumo";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Ind_Telemedida";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Cod_Factura_Sustituida";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Cod_Factura_Rectificada";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Codigo_Estado_Factura";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Tarifa";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Tipo_Consumo";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Tipo_Telegestion";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Contrato_Ext_PS";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Contrato_PS";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Sec_Factura";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Distribuidora";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Consumo_Tot_Act";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Consumo_Tot_React";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_desde_AE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_hasta_AE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_desde_R1";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_hasta_R1";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_desde_PFac";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_hasta_PFac";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_desde_Curva";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_hasta_Curva";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "CUPS_EXT";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Cod_Carga_ODS";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_act_ODS";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_act_DMCO";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Fecha_Recepcion";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = "Cod_Carga";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                PintaRecuadro(excelPackage, f, c); c++;


                #endregion

                var allCells = workSheet.Cells[1, 1, 1, 95];
                var cellFont = allCells.Style.Font;
                cellFont.Bold = true;

                if(lista_facturas != null)
                    foreach (EndesaEntity.medida.Peajes_Vista p in lista_facturas)
                {
                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = p.cups13; c++;
                    workSheet.Cells[f, c].Value = p.cups20; c++;
                    workSheet.Cells[f, c].Value = p.cod_factura; c++;
                    workSheet.Cells[f, c].Value = p.estado_factura; c++;
                    workSheet.Cells[f, c].Value = p.tipo_factura; c++;

                    workSheet.Cells[f, c].Value = p.fecha_factura;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                    workSheet.Cells[f, c].Value = p.fecha_desde;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                    workSheet.Cells[f, c].Value = p.fecha_hasta;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                    workSheet.Cells[f, c].Value = p.ind_cuadre_periodo_c; c++;

                    workSheet.Cells[f, c].Value = p.ffm_ae_total_c;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = p.ffm_r1_total_c;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = p.ffm_pmax_total_c;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = p.ffm_ae_total_c;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = p.fff_r1_total_c;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = p.fff_pfac_total_c;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = p.fff_imp_exc_pot_total_c;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    workSheet.Cells[f, c].Value = p.fff_imp_exc_pot;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    workSheet.Cells[f, c].Value = p.fff_imp_exc_r1;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    for (int i = 1; i <= 6; i++)
                    {
                        workSheet.Cells[f, c].Value = p.ffm_ae[i - 1];
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    }

                    for (int i = 1; i <= 6; i++)
                    {
                        workSheet.Cells[f, c].Value = p.ffm_r1[i - 1];
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    }

                    for (int i = 1; i <= 6; i++)
                    {
                        workSheet.Cells[f, c].Value = p.ffm_pmax[i - 1];
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    }

                    for (int i = 1; i <= 6; i++)
                    {
                        workSheet.Cells[f, c].Value = p.fff_ae[i - 1];
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    }

                    for (int i = 1; i <= 6; i++)
                    {
                        workSheet.Cells[f, c].Value = p.fff_r1[i - 1];
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    }

                    workSheet.Cells[f, c].Value = p.fff_r4_6;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;


                    for (int i = 1; i <= 6; i++)
                    {
                        workSheet.Cells[f, c].Value = p.fff_pfac[i - 1];
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    }

                    for (int i = 1; i <= 6; i++)
                    {
                        workSheet.Cells[f, c].Value = p.fff_imp_exc_pot1[i - 1];
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }

                    workSheet.Cells[f, c].Value = p.cups20_metra; c++;
                    workSheet.Cells[f, c].Value = p.procedencia; c++;
                    workSheet.Cells[f, c].Value = p.ind_perdidas; c++;

                    workSheet.Cells[f, c].Value = p.porcentaje_perdidas;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    workSheet.Cells[f, c].Value = p.pot_trafo_perdidas_va;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    workSheet.Cells[f, c].Value = p.tipo_pm; c++;
                    workSheet.Cells[f, c].Value = p.ind_autoconsumo; c++;
                    workSheet.Cells[f, c].Value = p.ind_telemedida; c++;
                    workSheet.Cells[f, c].Value = p.cod_factura_sustituida; c++;
                    workSheet.Cells[f, c].Value = p.cod_factura_rectificada; c++;
                    workSheet.Cells[f, c].Value = p.codigo_estado_factura; c++;
                    workSheet.Cells[f, c].Value = p.tarifa; c++;
                    workSheet.Cells[f, c].Value = p.tipo_consumo; c++;
                    workSheet.Cells[f, c].Value = p.tipo_telegestion; c++;
                    workSheet.Cells[f, c].Value = p.contrato_ext_ps; c++;
                    workSheet.Cells[f, c].Value = p.contrato_ps; c++;

                    workSheet.Cells[f, c].Value = p.sec_factura; c++;
                    workSheet.Cells[f, c].Value = p.distribuidora; c++;

                                            
                    workSheet.Cells[f, c].Value = p.consumo_tot_act;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    workSheet.Cells[f, c].Value = p.consumo_tot_react;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;


                    if(p.fecha_desde_ae > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_desde_ae;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; 
                    }
                    c++;

                    if (p.fecha_hasta_ae > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_hasta_ae;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.fecha_desde_r1 > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_desde_r1;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.fecha_hasta_r1 > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_hasta_r1;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.fecha_desde_pfac > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_desde_pfac;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.fecha_hasta_pfac > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_hasta_pfac;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.fecha_desde_curva > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_desde_curva;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.fecha_hasta_curva > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_hasta_curva;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    workSheet.Cells[f, c].Value = p.cups_ext; c++;

                    workSheet.Cells[f, c].Value = p.cod_carga_ods;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    if(p.fecha_act_ods > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_act_ods;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.fecha_act_dmco > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_act_dmco;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.FullDateTimePattern;
                    }
                    c++;

                    if (p.fecha_recepcion > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.fecha_recepcion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    workSheet.Cells[f, c].Value = p.cod_carga;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;


                }

                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:CQ1"].AutoFilter = true;
                allCells = workSheet.Cells[1, 1, f, c];
                allCells.AutoFitColumns();
                excelPackage.Save();

                return "";

            }catch(Exception ex)
            {
                return null;
            }

        }


        private void RecorreQuery_cups13(DateTime fd, List<string> lista_cups13)
        {
            StringBuilder sb = new StringBuilder();
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;

            MySQLDB dbmy;
            MySqlCommand commandmy;
            MySqlDataReader rmy;
            string strSql = "";

            bool firstOnly = true;
            int j = 0;
            int k = 0;
            int totalRegistros = 0;
            string o;

            try
            {


                // Borramos table tmp               

                strSql = "delete from t_ed_h_facts_atr_ml_tmp";
                Console.WriteLine("Ejecutando " + strSql);
                ficheroLog.Add("Ejecutando " + strSql);
                dbmy = new MySQLDB(MySQLDB.Esquemas.MED);
                commandmy = new MySqlCommand(strSql, dbmy.con);
                commandmy.ExecuteNonQuery();
                dbmy.CloseConnection();



                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(Consulta_Peajes(fd, lista_cups13), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    Console.CursorLeft = 0;
                    Console.Write("-");
                    Console.CursorLeft = 0;
                    Console.Write("\\");
                    Console.CursorLeft = 0;
                    Console.Write("|");
                    Console.CursorLeft = 0;
                    Console.Write("/");
                    Console.CursorLeft = 0;



                    j++;
                    k++;
                    if (firstOnly)
                    {
                        sb = null;
                        sb = new StringBuilder();
                        sb.Append("REPLACE INTO t_ed_h_facts_atr_ml_tmp");
                        sb.Append(" (cl_crto_ext_ps, id_crto_ext_ps, id_crto_ps, cd_sec_fact, cd_linea_neg, de_linea_neg, cd_empr_distr,");
                        sb.Append("cd_est_fact, cd_mes, id_fact, fh_fact, fh_ini_fact, fh_fin_fact, nm_importe, nm_meses, cd_tp_fact, de_tp_fact,");
                        sb.Append("id_fact_sust, cd_cups, cd_cups20, cd_cups_ext, cd_empr_distr_cne, de_empr_distr_cne, nm_med_potencia_1, nm_prec_potencia_1,");
                        sb.Append("nm_med_potencia_2, nm_prec_potencia_2, nm_med_potencia_3, nm_prec_potencia_3, nm_med_potencia_4, nm_prec_potencia_4,");
                        sb.Append("nm_med_potencia_5, nm_prec_potencia_5, nm_med_potencia_6, nm_prec_potencia_6, nm_med_potencia_7, nm_prec_potencia_7,");
                        sb.Append("nm_med_potencia_8, nm_prec_potencia_8, nm_med_potencia_9, nm_prec_potencia_9, nm_med_potencia_10, nm_prec_potencia_10,");
                        sb.Append("nm_med_activa_1, nm_prec_activa_1, nm_med_activa_2, nm_prec_activa_2, nm_med_activa_3, nm_prec_activa_3, nm_med_activa_4,");
                        sb.Append("nm_prec_activa_4, nm_med_activa_5, nm_prec_activa_5, nm_med_activa_6, nm_prec_activa_6, nm_med_activa_7, nm_prec_activa_7,");
                        sb.Append("nm_med_activa_8, nm_prec_activa_8, nm_med_activa_9, nm_prec_activa_9, nm_med_activa_10, nm_prec_activa_10, nm_med_reactiva_1,");
                        sb.Append("nm_prec_reactiva_1, nm_med_reactiva_2, nm_prec_reactiva_2, nm_med_reactiva_3, nm_prec_reactiva_3, nm_med_reactiva_4,");
                        sb.Append("nm_prec_reactiva_4, nm_med_reactiva_5, nm_prec_reactiva_5, nm_med_reactiva_6, nm_prec_reactiva_6, nm_med_reactiva_7,");
                        sb.Append("nm_prec_reactiva_7, nm_med_reactiva_8, nm_prec_reactiva_8, nm_med_reactiva_9, nm_prec_reactiva_9, nm_med_reactiva_10,");
                        sb.Append("nm_prec_reactiva_10, cd_concepto_1, cd_concepto_1_sce, de_concepto_1, im_concepto_1, cd_concepto_2, cd_concepto_2_sce,");
                        sb.Append("de_concepto_2, im_concepto_2, cd_concepto_3, cd_concepto_3_sce, de_concepto_3, im_concepto_3, cd_concepto_4, cd_concepto_4_sce,");
                        sb.Append("de_concepto_4, im_concepto_4, cd_concepto_5, cd_concepto_5_sce, de_concepto_5, im_concepto_5, cd_concepto_6, cd_concepto_6_sce,");
                        sb.Append("de_concepto_6, im_concepto_6, cd_concepto_7, cd_concepto_7_sce, de_concepto_7, im_concepto_7, cd_concepto_8, cd_concepto_8_sce,");
                        sb.Append("de_concepto_8, im_concepto_8, cd_concepto_9, cd_concepto_9_sce, de_concepto_9, im_concepto_9, cd_concepto_10, cd_concepto_10_sce,");
                        sb.Append("de_concepto_10, im_concepto_10, cd_concepto_11, cd_concepto_11_sce, de_concepto_11, im_concepto_11, cd_concepto_12, cd_concepto_12_sce,");
                        sb.Append("de_concepto_12, im_concepto_12, cd_concepto_13, cd_concepto_13_sce, de_concepto_13, im_concepto_13, cd_concepto_14, cd_concepto_14_sce,");
                        sb.Append("de_concepto_14, im_concepto_14, cd_concepto_15, cd_concepto_15_sce, de_concepto_15, im_concepto_15, cd_concepto_16, cd_concepto_16_sce,");
                        sb.Append("de_concepto_16, im_concepto_16, cd_concepto_17, cd_concepto_17_sce, de_concepto_17, im_concepto_17, cd_concepto_18, cd_concepto_18_sce,");
                        sb.Append("de_concepto_18, im_concepto_18, cd_concepto_19, cd_concepto_19_sce, de_concepto_19, im_concepto_19, cd_concepto_20, cd_concepto_20_sce,");
                        sb.Append("de_concepto_20, im_concepto_20, de_impuesto_1, nm_porcentaje_1, nm_base_1, nm_importe_1, de_impuesto_2, nm_porcentaje_2, nm_base_2, nm_importe_2,");
                        sb.Append("de_impuesto_3, nm_porcentaje_3, nm_base_3, nm_importe_3, de_impuesto_4, nm_porcentaje_4, nm_base_4, nm_importe_4, de_impuesto_5, nm_porcentaje_5,");
                        sb.Append("nm_base_5, nm_importe_5, cod_carga_ods, fh_act_ods, fh_act_dmco, cd_tarifa, cd_tarifa_ff, nm_pot_ctatada, cd_comercializadora, cd_tipo_consumo,");
                        sb.Append("cd_municipio, cd_cups20_metra, cd_tp_teleg, consumo_tot_act, consumo_punta, consumo_llano, consumo_valle, consumo_activa4, consumo_activa5,");
                        sb.Append("consumo_activa6, consumo_tot_react, consumo_reactiva1, consumo_reactiva2, consumo_reactiva3, consumo_reactiva4, consumo_reactiva5, consumo_reactiva6,");
                        sb.Append("potencia_maxpunta, potencia_maxllano, potencia_maxvalle, potencia_max4, potencia_max5, potencia_max6, consumo_supervalle, fh_recepcion, cod_carga,");
                        sb.Append("id_pseudofactura, cd_comercializadora_fact, cd_empr_comer_cne, id_fact_recti, id_nif_dist, de_comcnmc, im_salfact, tp_moneda, fh_feboe, lg_med_perdida,");
                        sb.Append("nm_vastrafo, nm_porc_perdida, cd_fact_curva, fh_desde_cch, fh_hasta_cch, nm_num_mes, fh_lect_ant_punta, cd_tp_fuente_ant_punta, nm_lect_ant_punta,");
                        sb.Append("fh_lect_act_punta, cd_tp_fuente_act_punta, nm_lect_act_punta, im_ajuste_integr_punta, cd_tp_anoma_punta, fh_lect_ant_llano, cd_tp_fuente_ant_llano,");
                        sb.Append("nm_lect_ant_llano, fh_lect_act_llano, cd_tp_fuente_act_llano, nm_lect_act_llano, im_ajuste_integr_llano, cd_tp_anoma_llano, fh_lect_ant_valle,");
                        sb.Append("cd_tp_fuente_ant_valle, nm_lect_ant_valle, fh_lect_act_valle, cd_tp_fuente_act_valle, nm_lect_act_valle, im_ajuste_integr_valle, cd_tp_anoma_valle,");
                        sb.Append("fh_lect_ant_activa4, cd_tp_fuente_ant_activa4, nm_lect_ant_activa4, fh_lect_act_activa4, cd_tp_fuente_act_activa4, nm_lect_act_activa4,");
                        sb.Append("im_ajuste_integr_activa4, cd_tp_anoma_activa4, fh_lect_ant_activa5, cd_tp_fuente_ant_activa5, nm_lect_ant_activa5, fh_lect_act_activa5,");
                        sb.Append("cd_tp_fuente_act_activa5, nm_lect_act_activa5, im_ajuste_integr_activa5, cd_tp_anoma_activa5, fh_lect_ant_activa6, cd_tp_fuente_ant_activa6,");
                        sb.Append("nm_lect_ant_activa6, fh_lect_act_activa6, cd_tp_fuente_act_activa6, nm_lect_act_activa6, im_ajuste_integr_activa6, cd_tp_anoma_activa6,");
                        sb.Append("fh_lect_ant_reactiva1, cd_tp_fuente_ant_reactiva1, nm_lect_ant_reactiva1, fh_lect_act_reactiva1, cd_tp_fuente_act_reactiva1, nm_lect_act_reactiva1,");
                        sb.Append("im_ajuste_integr_reactiva1, cd_tp_anoma_reactiva1, fh_lect_ant_reactiva2, cd_tp_fuente_ant_reactiva2, nm_lect_ant_reactiva2, fh_lect_act_reactiva2,");
                        sb.Append("cd_tp_fuente_act_reactiva2, nm_lect_act_reactiva2, im_ajuste_integr_reactiva2, cd_tp_anoma_reactiva2, fh_lect_ant_reactiva3, cd_tp_fuente_ant_reactiva3,");
                        sb.Append("nm_lect_ant_reactiva3, fh_lect_act_reactiva3, cd_tp_fuente_act_reactiva3, nm_lect_act_reactiva3, im_ajuste_integr_reactiva3, cd_tp_anoma_reactiva3,");
                        sb.Append("fh_lect_ant_reactiva4, cd_tp_fuente_ant_reactiva4, nm_lect_ant_reactiva4, fh_lect_act_reactiva4, cd_tp_fuente_act_reactiva4, nm_lect_act_reactiva4,");
                        sb.Append("im_ajuste_integr_reactiva4, cd_tp_anoma_reactiva4, fh_lect_ant_reactiva5, cd_tp_fuente_ant_reactiva5, nm_lect_ant_reactiva5, fh_lect_act_reactiva5,");
                        sb.Append("cd_tp_fuente_act_reactiva5, nm_lect_act_reactiva5, im_ajuste_integr_reactiva5, cd_tp_anoma_reactiva5, fh_lect_ant_reactiva6, cd_tp_fuente_ant_reactiva6,");
                        sb.Append("nm_lect_ant_reactiva6, fh_lect_act_reactiva6, cd_tp_fuente_act_reactiva6, nm_lect_act_reactiva6, im_ajuste_integr_reactiva6, cd_tp_anoma_reactiva6,");
                        sb.Append("fh_lect_ant_maxpunta, cd_tp_fuente_ant_maxpunta, nm_lect_ant_maxpunta, fh_lect_act_maxpunta, cd_tp_fuente_act_maxpunta, nm_lect_act_maxpunta,");
                        sb.Append("im_ajuste_integr_maxpunta, cd_tp_anoma_maxpunta, fh_lect_ant_maxllano, cd_tp_fuente_ant_maxllano, nm_lect_ant_maxllano, fh_lect_act_maxllano,");
                        sb.Append("cd_tp_fuente_act_maxllano, nm_lect_act_maxllano, im_ajuste_integr_maxllano, cd_tp_anoma_maxllano, fh_lect_ant_maxvalle, cd_tp_fuente_ant_maxvalle,");
                        sb.Append("nm_lect_ant_maxvalle, fh_lect_act_maxvalle, cd_tp_fuente_act_maxvalle, nm_lect_act_maxvalle, im_ajuste_integr_maxvalle, cd_tp_anoma_maxvalle,");
                        sb.Append("fh_lect_ant_max4, cd_tp_fuente_ant_max4, nm_lect_ant_max4, fh_lect_act_max4, cd_tp_fuente_act_max4, nm_lect_act_max4, im_ajuste_integr_max4,");
                        sb.Append("cd_tp_anoma_max4, fh_lect_ant_max5, cd_tp_fuente_ant_max5, nm_lect_ant_max5, fh_lect_act_max5, cd_tp_fuente_act_max5, nm_lect_act_max5,");
                        sb.Append("im_ajuste_integr_max5, cd_tp_anoma_max5, fh_lect_ant_max6, cd_tp_fuente_ant_max6, nm_lect_ant_max6, fh_lect_act_max6, cd_tp_fuente_act_max6,");
                        sb.Append("nm_lect_act_max6, im_ajuste_integr_max6, cd_tp_anoma_max6, fh_lect_ant_supervalle, cd_tp_fuente_ant_supervalle, nm_lect_ant_supervalle,");
                        sb.Append("fh_lect_act_supervalle, cd_tp_fuente_act_supervalle, nm_lect_act_supervalle, im_ajuste_integr_supervalle, cd_tp_anoma_supervalle,");
                        sb.Append("cd_concepto_reper_1, im_concepto_reper_1, cd_concepto_reper_2, im_concepto_reper_2, cd_concepto_reper_3, im_concepto_reper_3, cd_concepto_reper_4,");
                        sb.Append("im_concepto_reper_4, cd_concepto_reper_5, im_concepto_reper_5, cd_concepto_reper_6, im_concepto_reper_6, cd_concepto_reper_7, im_concepto_reper_7,");
                        sb.Append("cd_concepto_reper_8, im_concepto_reper_8, cd_concepto_reper_9, im_concepto_reper_9, cd_concepto_reper_10, im_concepto_reper_10, cd_concepto_reper_11,");
                        sb.Append("im_concepto_reper_11, cd_concepto_reper_12, im_concepto_reper_12, cd_cmunicip_atr, de_est_fact, fh_desde_ener, fh_hasta_ener, fh_desde_pot,");
                        sb.Append("fh_hasta_pot, cd_modo_crtl_pot, nm_dias_fact, im_tot_fact, im_saldo_tot_fact, nm_tot_recibos, fh_valor, fh_lim_pago, id_remesa, cd_munic_red,");
                        sb.Append("lg_penaliza_icp, identificador, cd_mensaje, cd_doc_ref, cd_tp_doc, de_tp_doc, de_tp_aceso_serv, lg_relevancia_recl_pdp, de_empr_comer_cne,");
                        sb.Append("cd_huella_fact_atr, fichero, cd_huella_blq_fich, fh_publicacion, fh_public_fichero, cd_solicitud, cd_sec_solicitud, lg_atr, cd_tp_reg, de_tp_reg,");
                        sb.Append("cd_expdte, lg_bloq_af, lg_recl_abta, cd_origen_fact, de_origen_fact, cd_aniofactura, cd_pais, lg_autoconsumo, cd_tp_pm, lg_duracioninfanio,");
                        sb.Append("fh_creacion, cd_usuario_creador, de_usuario_creador, fh_mod, cd_usuario_mod, de_usuario_mod, cd_ubicacion, de_ubicacion, id_exp_af, lg_registro_borrado,");
                        sb.Append("de_tipo_borrado, cd_rg_presion, de_rg_presion, cd_metodo_fact, de_metodo_fact, lg_telemedida, cd_tp_gasinera, de_tp_gasinera, fh_desde_reac, fh_hasta_reac,");
                        sb.Append("de_marca_back, lg_contrato_simul, cd_id_producto, cd_tp_producto, de_tp_producto, lg_arrastre_penali, nm_med_capacitiva, nm_prec_capacitiva, nm_exceso_pot_p1,");
                        sb.Append("nm_exceso_pot_p2, nm_exceso_pot_p3, nm_exceso_pot_p4, nm_exceso_pot_p5, nm_exceso_pot_p6, nm_consumo_medio_5a, nm_consumo_medio) values ");
                        firstOnly = false;
                    }

                    #region Campos
                    if (r["cl_crto_ext_ps"] != System.DBNull.Value)
                        sb.Append("('").Append(r["cl_crto_ext_ps"].ToString()).Append("',");
                    else
                        sb.Append("(null,");

                    if (r["id_crto_ext_ps"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_crto_ext_ps"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_crto_ps"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_crto_ps"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_sec_fact"] != System.DBNull.Value)
                        sb.Append(r["cd_sec_fact"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_linea_neg"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_linea_neg"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_linea_neg"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_linea_neg"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_empr_distr"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_empr_distr"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_est_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_est_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_mes"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_mes"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_fact"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_ini_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_ini_fact"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_fin_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_fin_fact"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_importe"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_importe"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_meses"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_meses"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_fact_sust"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_fact_sust"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_cups"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_cups"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_cups_ext"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_cups_ext"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_empr_distr_cne"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_empr_distr_cne"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_empr_distr_cne"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_empr_distr_cne"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    //nm_med_potencia_
                    for (int i = 0; i < 10; i++)
                    {
                        if (r["nm_med_potencia_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_med_potencia_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["nm_prec_potencia_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_prec_potencia_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    // nm_med_activa_
                    for (int i = 0; i < 10; i++)
                    {
                        if (r["nm_med_activa_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_med_activa_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["nm_prec_activa_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_prec_activa_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    //nm_med_reactiva_
                    for (int i = 0; i < 10; i++)
                    {
                        if (r["nm_med_reactiva_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_med_reactiva_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["nm_prec_reactiva_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_prec_reactiva_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    //cd_concepto_
                    for (int i = 0; i < 20; i++)
                    {
                        if (r["cd_concepto_" + (i + 1)] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_concepto_" + (i + 1)].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_concepto_" + (i + 1) + "_sce"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_concepto_" + (i + 1) + "_sce"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_concepto_" + (i + 1)] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_concepto_" + (i + 1)].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["im_concepto_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["im_concepto_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    // de_impuesto_1
                    for (int i = 0; i < 5; i++)
                    {
                        if (r["de_impuesto_" + (i + 1)] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_impuesto_" + (i + 1)].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_porcentaje_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_porcentaje_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["nm_base_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_base_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["nm_importe_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_importe_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    if (r["cod_carga_ods"] != System.DBNull.Value)
                        sb.Append(r["cod_carga_ods"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_act_ods"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_act_ods"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_act_dmco"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_act_dmco"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tarifa"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tarifa"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tarifa_ff"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tarifa_ff"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_pot_ctatada"] != System.DBNull.Value)
                        sb.Append(r["nm_pot_ctatada"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_comercializadora"] != System.DBNull.Value)
                        sb.Append(r["cd_comercializadora"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_tipo_consumo"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tipo_consumo"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_municipio"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tipo_consumo"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_cups20_metra"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_cups20_metra"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_teleg"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_teleg"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["consumo_tot_act"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_tot_act"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_punta"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_punta"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_llano"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_llano"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_valle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_valle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_activa4"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_activa4"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_activa5"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_activa5"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_activa6"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_activa6"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_tot_react"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_tot_react"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    for (int i = 0; i < 6; i++)
                    {
                        if (r["consumo_reactiva" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["consumo_reactiva" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    if (r["potencia_maxpunta"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["potencia_maxpunta"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["potencia_maxllano"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["potencia_maxllano"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["potencia_maxvalle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["potencia_maxvalle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["potencia_max4"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["potencia_max4"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["potencia_max5"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["potencia_max5"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["potencia_max6"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["potencia_max6"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_supervalle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_supervalle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_recepcion"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_recepcion"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cod_carga"] != System.DBNull.Value)
                        sb.Append(r["cod_carga"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["id_pseudofactura"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_pseudofactura"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_comercializadora_fact"] != System.DBNull.Value)
                        sb.Append(r["cd_comercializadora_fact"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_empr_comer_cne"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_empr_comer_cne"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_fact_recti"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_fact_recti"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_nif_dist"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_nif_dist"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_comcnmc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_comcnmc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["im_salfact"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_salfact"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["tp_moneda"] != System.DBNull.Value)
                        sb.Append("'").Append(r["tp_moneda"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_feboe"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_feboe"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_med_perdida"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_med_perdida"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_vastrafo"] != System.DBNull.Value)
                        sb.Append(r["nm_vastrafo"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_porc_perdida"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_porc_perdida"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_fact_curva"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_fact_curva"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_desde_cch"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_desde_cch"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_hasta_cch"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_hasta_cch"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_num_mes"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_num_mes"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_ant_punta"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_punta"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_punta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_punta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_punta"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_punta"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_act_punta"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_punta"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_punta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_punta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_punta"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_punta"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_punta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_punta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_punta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_punta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_ant_llano"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_llano"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_llano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_llano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_llano"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_llano"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_act_llano"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_llano"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_llano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_llano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_llano"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_llano"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_llano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_llano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_llano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_llano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_ant_valle"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_valle"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_valle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_valle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_valle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_valle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_act_valle"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_valle"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_valle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_valle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_valle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_valle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_valle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_valle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_valle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_valle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    for (int i = 4; i <= 6; i++)
                    {
                        if (r["fh_lect_ant_activa" + i] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_activa" + i]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_fuente_ant_activa" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_fuente_ant_activa" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_lect_ant_activa" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_lect_ant_activa" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["fh_lect_act_activa" + i] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_activa" + i]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_fuente_act_activa" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_fuente_act_activa" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_lect_act_activa" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_lect_act_activa" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["im_ajuste_integr_activa" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["im_ajuste_integr_activa" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_anoma_activa" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_anoma_activa" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");
                    }

                    // fh_lect_ant_reactiva
                    for (int i = 1; i <= 6; i++)
                    {
                        if (r["fh_lect_ant_reactiva" + i] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_reactiva" + i]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_fuente_ant_reactiva" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_fuente_ant_reactiva" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_lect_ant_reactiva" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_lect_ant_reactiva" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["fh_lect_act_reactiva" + i] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_reactiva" + i]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_fuente_act_reactiva" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_fuente_act_reactiva" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_lect_act_reactiva" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_lect_act_reactiva" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["im_ajuste_integr_reactiva" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["im_ajuste_integr_reactiva" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_anoma_reactiva" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_anoma_reactiva" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");
                    }


                    if (r["fh_lect_ant_maxpunta"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_maxpunta"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_maxpunta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_maxpunta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_maxpunta"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_maxpunta"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_act_maxpunta"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_maxpunta"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_maxpunta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_maxpunta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_maxpunta"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_maxpunta"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_maxpunta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_maxpunta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_maxpunta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_maxpunta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_ant_maxllano"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_maxllano"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_maxllano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_maxllano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_maxllano"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_maxllano"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");


                    if (r["fh_lect_act_maxllano"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_maxllano"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_maxllano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_maxllano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_maxllano"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_maxllano"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_maxllano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_maxllano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_maxllano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_maxllano"].ToString()).Append("',");
                    else
                        sb.Append("null,");


                    if (r["fh_lect_ant_maxvalle"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_maxvalle"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_maxvalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_maxvalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_maxvalle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_maxvalle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_act_maxvalle"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_maxvalle"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_maxvalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_maxvalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_maxvalle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_maxvalle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_maxvalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_maxvalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_maxvalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_maxvalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    // fh_lect_ant_max4
                    for (int i = 4; i <= 6; i++)
                    {
                        if (r["fh_lect_ant_max" + i] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_max" + i]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_fuente_ant_max" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_fuente_ant_max" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_lect_ant_max" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_lect_ant_max" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["fh_lect_act_max" + i] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_max" + i]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_fuente_act_max" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_fuente_act_max" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_lect_act_max" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_lect_act_max" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["im_ajuste_integr_max" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["im_ajuste_integr_max" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_anoma_max" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_anoma_max" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");
                    }

                    if (r["fh_lect_ant_supervalle"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_supervalle"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_supervalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_supervalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_supervalle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_supervalle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_act_supervalle"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_supervalle"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_supervalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_supervalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_supervalle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_supervalle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_supervalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_supervalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_supervalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_supervalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");


                    // cd_concepto_reper_
                    for (int i = 1; i <= 12; i++)
                    {
                        if (r["cd_concepto_reper_" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_concepto_reper_" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["im_concepto_reper_" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["im_concepto_reper_" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }


                    if (r["cd_cmunicip_atr"] != System.DBNull.Value)
                        sb.Append(r["cd_cmunicip_atr"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["de_est_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_est_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_desde_ener"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_desde_ener"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_hasta_ener"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_hasta_ener"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_desde_pot"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_desde_pot"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_hasta_pot"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_hasta_pot"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");


                    if (r["cd_modo_crtl_pot"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_modo_crtl_pot"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_dias_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["nm_dias_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["im_tot_fact"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_tot_fact"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_saldo_tot_fact"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_saldo_tot_fact"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_tot_recibos"] != System.DBNull.Value)
                        sb.Append(r["nm_tot_recibos"].ToString()).Append(",");
                    else
                        sb.Append("null,");


                    if (r["fh_valor"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_valor"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_lim_pago"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lim_pago"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");


                    if (r["id_remesa"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_remesa"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_munic_red"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_munic_red"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_penaliza_icp"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_penaliza_icp"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["identificador"] != System.DBNull.Value)
                        sb.Append("'").Append(r["identificador"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_mensaje"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_mensaje"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_doc_ref"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_doc_ref"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_doc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_doc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_doc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_doc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_aceso_serv"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_aceso_serv"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_relevancia_recl_pdp"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_relevancia_recl_pdp"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_empr_comer_cne"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_empr_comer_cne"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_huella_fact_atr"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_huella_fact_atr"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fichero"] != System.DBNull.Value)
                        sb.Append("'").Append(r["fichero"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_huella_blq_fich"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_huella_blq_fich"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_publicacion"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_publicacion"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_public_fichero"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_public_fichero"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_solicitud"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_solicitud"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_sec_solicitud"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_sec_solicitud"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_atr"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_atr"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_reg"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_reg"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_reg"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_reg"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_expdte"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_expdte"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_bloq_af"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_bloq_af"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_recl_abta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_recl_abta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_origen_fact"] != System.DBNull.Value)
                        sb.Append(r["cd_origen_fact"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["de_origen_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_origen_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_aniofactura"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_aniofactura"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_pais"] != System.DBNull.Value)
                        sb.Append(r["cd_pais"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["lg_autoconsumo"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_autoconsumo"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_pm"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_pm"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_duracioninfanio"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_duracioninfanio"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_creacion"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_creacion"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_usuario_creador"] != System.DBNull.Value)
                        sb.Append(r["cd_usuario_creador"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["de_usuario_creador"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_usuario_creador"].ToString()).Append("',");
                    else
                        sb.Append("null,");


                    if (r["fh_mod"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_mod"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_usuario_mod"] != System.DBNull.Value)
                        sb.Append(r["cd_usuario_mod"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["de_usuario_mod"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_usuario_mod"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_ubicacion"] != System.DBNull.Value)
                        sb.Append(r["cd_ubicacion"].ToString()).Append(",");
                    else
                        sb.Append("null,");


                    if (r["de_ubicacion"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_ubicacion"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_exp_af"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_exp_af"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_registro_borrado"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_registro_borrado"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tipo_borrado"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tipo_borrado"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_rg_presion"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_rg_presion"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_rg_presion"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_rg_presion"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_metodo_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_metodo_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_metodo_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_metodo_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_telemedida"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_telemedida"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_gasinera"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_gasinera"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_gasinera"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_gasinera"].ToString()).Append("',");
                    else
                        sb.Append("null,");


                    if (r["fh_desde_reac"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_desde_reac"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_hasta_reac"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_hasta_reac"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_marca_back"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_marca_back"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_contrato_simul"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_contrato_simul"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_id_producto"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_id_producto"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_producto"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_producto"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_producto"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_producto"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_arrastre_penali"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_arrastre_penali"].ToString()).Append("',");
                    else
                        sb.Append("null,");




                    if (r["nm_med_capacitiva"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_med_capacitiva"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_prec_capacitiva"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_prec_capacitiva"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");



                    for (int i = 1; i <= 6; i++)
                    {
                        if (r["nm_exceso_pot_p" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_exceso_pot_p" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    if (r["nm_consumo_medio_5a"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_consumo_medio_5a"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_consumo_medio"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_consumo_medio"]).ToString("").Replace(",", ".")).Append("),");
                    else
                        sb.Append("null),");


                    #endregion


                    if (j == 100)
                    {
                        Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                        firstOnly = true;
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
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
                    dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                    commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                    commandmy.ExecuteNonQuery();
                    dbmy.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    j = 0;
                }

                // Construimos conceptos
                for (int i = 1; i < 21; i++)
                {
                    strSql = "REPLACE INTO t_ed_h_facts_atr_ml_conceptos"
                        + " SELECT f.id_fact, f.cd_cups_ext,"
                        + " f.cd_concepto_" + i + ","
                        + " f.cd_concepto_" + i + "_sce,"
                        + " f.de_concepto_" + i + ","
                        + " f.im_concepto_" + i
                        + " FROM t_ed_h_facts_atr_ml_tmp f"
                        + " where f.cd_concepto_" + i + " <> ''";                        
                    Console.WriteLine("Ejecutando " + strSql);
                    ficheroLog.Add("Ejecutando " + strSql);
                    dbmy = new MySQLDB(MySQLDB.Esquemas.MED);
                    commandmy = new MySqlCommand(strSql, dbmy.con);
                    commandmy.ExecuteNonQuery();
                    dbmy.CloseConnection();
                }

                // Volcamos tmp sobre tabla normal

                strSql = "replace into t_ed_h_facts_atr_ml select * from t_ed_h_facts_atr_ml_tmp";
                Console.WriteLine("Ejecutando " + strSql);
                ficheroLog.Add("Ejecutando " + strSql);
                dbmy = new MySQLDB(MySQLDB.Esquemas.MED);
                commandmy = new MySqlCommand(strSql, dbmy.con);
                commandmy.ExecuteNonQuery();
                dbmy.CloseConnection();


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                ficheroLog.addError(ex.Message);
            }
        }

        private void RecorreQuery_cups20(DateTime fd, Dictionary<string, string> dic_cups20)
        {
            StringBuilder sb = new StringBuilder();
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;

            MySQLDB dbmy;
            MySqlCommand commandmy;
            MySqlDataReader rmy;
            string strSql = "";

            bool firstOnly = true;
            int j = 0;
            int k = 0;
            int totalRegistros = 0;
            string o;

            try
            {


                // Borramos table tmp               

                strSql = "delete from t_ed_h_facts_atr_ml_tmp";
                Console.WriteLine("Ejecutando " + strSql);
                ficheroLog.Add("Ejecutando " + strSql);
                dbmy = new MySQLDB(MySQLDB.Esquemas.MED);
                commandmy = new MySqlCommand(strSql, dbmy.con);
                commandmy.ExecuteNonQuery();
                dbmy.CloseConnection();



                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                Console.WriteLine("Ejecutando Consulta BI" );
                command = new OdbcCommand(Consulta_Peajes_cups20(fd, dic_cups20), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    Console.CursorLeft = 0;
                    Console.Write("-");
                    Console.CursorLeft = 0;
                    Console.Write("\\");
                    Console.CursorLeft = 0;
                    Console.Write("|");
                    Console.CursorLeft = 0;
                    Console.Write("/");
                    Console.CursorLeft = 0;



                    j++;
                    k++;
                    if (firstOnly)
                    {
                        sb = null;
                        sb = new StringBuilder();
                        sb.Append("REPLACE INTO t_ed_h_facts_atr_ml_tmp");
                        sb.Append(" (cl_crto_ext_ps, id_crto_ext_ps, id_crto_ps, cd_sec_fact, cd_linea_neg, de_linea_neg, cd_empr_distr,");
                        sb.Append("cd_est_fact, cd_mes, id_fact, fh_fact, fh_ini_fact, fh_fin_fact, nm_importe, nm_meses, cd_tp_fact, de_tp_fact,");
                        sb.Append("id_fact_sust, cd_cups, cd_cups20, cd_cups_ext, cd_empr_distr_cne, de_empr_distr_cne, nm_med_potencia_1, nm_prec_potencia_1,");
                        sb.Append("nm_med_potencia_2, nm_prec_potencia_2, nm_med_potencia_3, nm_prec_potencia_3, nm_med_potencia_4, nm_prec_potencia_4,");
                        sb.Append("nm_med_potencia_5, nm_prec_potencia_5, nm_med_potencia_6, nm_prec_potencia_6, nm_med_potencia_7, nm_prec_potencia_7,");
                        sb.Append("nm_med_potencia_8, nm_prec_potencia_8, nm_med_potencia_9, nm_prec_potencia_9, nm_med_potencia_10, nm_prec_potencia_10,");
                        sb.Append("nm_med_activa_1, nm_prec_activa_1, nm_med_activa_2, nm_prec_activa_2, nm_med_activa_3, nm_prec_activa_3, nm_med_activa_4,");
                        sb.Append("nm_prec_activa_4, nm_med_activa_5, nm_prec_activa_5, nm_med_activa_6, nm_prec_activa_6, nm_med_activa_7, nm_prec_activa_7,");
                        sb.Append("nm_med_activa_8, nm_prec_activa_8, nm_med_activa_9, nm_prec_activa_9, nm_med_activa_10, nm_prec_activa_10, nm_med_reactiva_1,");
                        sb.Append("nm_prec_reactiva_1, nm_med_reactiva_2, nm_prec_reactiva_2, nm_med_reactiva_3, nm_prec_reactiva_3, nm_med_reactiva_4,");
                        sb.Append("nm_prec_reactiva_4, nm_med_reactiva_5, nm_prec_reactiva_5, nm_med_reactiva_6, nm_prec_reactiva_6, nm_med_reactiva_7,");
                        sb.Append("nm_prec_reactiva_7, nm_med_reactiva_8, nm_prec_reactiva_8, nm_med_reactiva_9, nm_prec_reactiva_9, nm_med_reactiva_10,");
                        sb.Append("nm_prec_reactiva_10, cd_concepto_1, cd_concepto_1_sce, de_concepto_1, im_concepto_1, cd_concepto_2, cd_concepto_2_sce,");
                        sb.Append("de_concepto_2, im_concepto_2, cd_concepto_3, cd_concepto_3_sce, de_concepto_3, im_concepto_3, cd_concepto_4, cd_concepto_4_sce,");
                        sb.Append("de_concepto_4, im_concepto_4, cd_concepto_5, cd_concepto_5_sce, de_concepto_5, im_concepto_5, cd_concepto_6, cd_concepto_6_sce,");
                        sb.Append("de_concepto_6, im_concepto_6, cd_concepto_7, cd_concepto_7_sce, de_concepto_7, im_concepto_7, cd_concepto_8, cd_concepto_8_sce,");
                        sb.Append("de_concepto_8, im_concepto_8, cd_concepto_9, cd_concepto_9_sce, de_concepto_9, im_concepto_9, cd_concepto_10, cd_concepto_10_sce,");
                        sb.Append("de_concepto_10, im_concepto_10, cd_concepto_11, cd_concepto_11_sce, de_concepto_11, im_concepto_11, cd_concepto_12, cd_concepto_12_sce,");
                        sb.Append("de_concepto_12, im_concepto_12, cd_concepto_13, cd_concepto_13_sce, de_concepto_13, im_concepto_13, cd_concepto_14, cd_concepto_14_sce,");
                        sb.Append("de_concepto_14, im_concepto_14, cd_concepto_15, cd_concepto_15_sce, de_concepto_15, im_concepto_15, cd_concepto_16, cd_concepto_16_sce,");
                        sb.Append("de_concepto_16, im_concepto_16, cd_concepto_17, cd_concepto_17_sce, de_concepto_17, im_concepto_17, cd_concepto_18, cd_concepto_18_sce,");
                        sb.Append("de_concepto_18, im_concepto_18, cd_concepto_19, cd_concepto_19_sce, de_concepto_19, im_concepto_19, cd_concepto_20, cd_concepto_20_sce,");
                        sb.Append("de_concepto_20, im_concepto_20, de_impuesto_1, nm_porcentaje_1, nm_base_1, nm_importe_1, de_impuesto_2, nm_porcentaje_2, nm_base_2, nm_importe_2,");
                        sb.Append("de_impuesto_3, nm_porcentaje_3, nm_base_3, nm_importe_3, de_impuesto_4, nm_porcentaje_4, nm_base_4, nm_importe_4, de_impuesto_5, nm_porcentaje_5,");
                        sb.Append("nm_base_5, nm_importe_5, cod_carga_ods, fh_act_ods, fh_act_dmco, cd_tarifa, cd_tarifa_ff, nm_pot_ctatada, cd_comercializadora, cd_tipo_consumo,");
                        sb.Append("cd_municipio, cd_cups20_metra, cd_tp_teleg, consumo_tot_act, consumo_punta, consumo_llano, consumo_valle, consumo_activa4, consumo_activa5,");
                        sb.Append("consumo_activa6, consumo_tot_react, consumo_reactiva1, consumo_reactiva2, consumo_reactiva3, consumo_reactiva4, consumo_reactiva5, consumo_reactiva6,");
                        sb.Append("potencia_maxpunta, potencia_maxllano, potencia_maxvalle, potencia_max4, potencia_max5, potencia_max6, consumo_supervalle, fh_recepcion, cod_carga,");
                        sb.Append("id_pseudofactura, cd_comercializadora_fact, cd_empr_comer_cne, id_fact_recti, id_nif_dist, de_comcnmc, im_salfact, tp_moneda, fh_feboe, lg_med_perdida,");
                        sb.Append("nm_vastrafo, nm_porc_perdida, cd_fact_curva, fh_desde_cch, fh_hasta_cch, nm_num_mes, fh_lect_ant_punta, cd_tp_fuente_ant_punta, nm_lect_ant_punta,");
                        sb.Append("fh_lect_act_punta, cd_tp_fuente_act_punta, nm_lect_act_punta, im_ajuste_integr_punta, cd_tp_anoma_punta, fh_lect_ant_llano, cd_tp_fuente_ant_llano,");
                        sb.Append("nm_lect_ant_llano, fh_lect_act_llano, cd_tp_fuente_act_llano, nm_lect_act_llano, im_ajuste_integr_llano, cd_tp_anoma_llano, fh_lect_ant_valle,");
                        sb.Append("cd_tp_fuente_ant_valle, nm_lect_ant_valle, fh_lect_act_valle, cd_tp_fuente_act_valle, nm_lect_act_valle, im_ajuste_integr_valle, cd_tp_anoma_valle,");
                        sb.Append("fh_lect_ant_activa4, cd_tp_fuente_ant_activa4, nm_lect_ant_activa4, fh_lect_act_activa4, cd_tp_fuente_act_activa4, nm_lect_act_activa4,");
                        sb.Append("im_ajuste_integr_activa4, cd_tp_anoma_activa4, fh_lect_ant_activa5, cd_tp_fuente_ant_activa5, nm_lect_ant_activa5, fh_lect_act_activa5,");
                        sb.Append("cd_tp_fuente_act_activa5, nm_lect_act_activa5, im_ajuste_integr_activa5, cd_tp_anoma_activa5, fh_lect_ant_activa6, cd_tp_fuente_ant_activa6,");
                        sb.Append("nm_lect_ant_activa6, fh_lect_act_activa6, cd_tp_fuente_act_activa6, nm_lect_act_activa6, im_ajuste_integr_activa6, cd_tp_anoma_activa6,");
                        sb.Append("fh_lect_ant_reactiva1, cd_tp_fuente_ant_reactiva1, nm_lect_ant_reactiva1, fh_lect_act_reactiva1, cd_tp_fuente_act_reactiva1, nm_lect_act_reactiva1,");
                        sb.Append("im_ajuste_integr_reactiva1, cd_tp_anoma_reactiva1, fh_lect_ant_reactiva2, cd_tp_fuente_ant_reactiva2, nm_lect_ant_reactiva2, fh_lect_act_reactiva2,");
                        sb.Append("cd_tp_fuente_act_reactiva2, nm_lect_act_reactiva2, im_ajuste_integr_reactiva2, cd_tp_anoma_reactiva2, fh_lect_ant_reactiva3, cd_tp_fuente_ant_reactiva3,");
                        sb.Append("nm_lect_ant_reactiva3, fh_lect_act_reactiva3, cd_tp_fuente_act_reactiva3, nm_lect_act_reactiva3, im_ajuste_integr_reactiva3, cd_tp_anoma_reactiva3,");
                        sb.Append("fh_lect_ant_reactiva4, cd_tp_fuente_ant_reactiva4, nm_lect_ant_reactiva4, fh_lect_act_reactiva4, cd_tp_fuente_act_reactiva4, nm_lect_act_reactiva4,");
                        sb.Append("im_ajuste_integr_reactiva4, cd_tp_anoma_reactiva4, fh_lect_ant_reactiva5, cd_tp_fuente_ant_reactiva5, nm_lect_ant_reactiva5, fh_lect_act_reactiva5,");
                        sb.Append("cd_tp_fuente_act_reactiva5, nm_lect_act_reactiva5, im_ajuste_integr_reactiva5, cd_tp_anoma_reactiva5, fh_lect_ant_reactiva6, cd_tp_fuente_ant_reactiva6,");
                        sb.Append("nm_lect_ant_reactiva6, fh_lect_act_reactiva6, cd_tp_fuente_act_reactiva6, nm_lect_act_reactiva6, im_ajuste_integr_reactiva6, cd_tp_anoma_reactiva6,");
                        sb.Append("fh_lect_ant_maxpunta, cd_tp_fuente_ant_maxpunta, nm_lect_ant_maxpunta, fh_lect_act_maxpunta, cd_tp_fuente_act_maxpunta, nm_lect_act_maxpunta,");
                        sb.Append("im_ajuste_integr_maxpunta, cd_tp_anoma_maxpunta, fh_lect_ant_maxllano, cd_tp_fuente_ant_maxllano, nm_lect_ant_maxllano, fh_lect_act_maxllano,");
                        sb.Append("cd_tp_fuente_act_maxllano, nm_lect_act_maxllano, im_ajuste_integr_maxllano, cd_tp_anoma_maxllano, fh_lect_ant_maxvalle, cd_tp_fuente_ant_maxvalle,");
                        sb.Append("nm_lect_ant_maxvalle, fh_lect_act_maxvalle, cd_tp_fuente_act_maxvalle, nm_lect_act_maxvalle, im_ajuste_integr_maxvalle, cd_tp_anoma_maxvalle,");
                        sb.Append("fh_lect_ant_max4, cd_tp_fuente_ant_max4, nm_lect_ant_max4, fh_lect_act_max4, cd_tp_fuente_act_max4, nm_lect_act_max4, im_ajuste_integr_max4,");
                        sb.Append("cd_tp_anoma_max4, fh_lect_ant_max5, cd_tp_fuente_ant_max5, nm_lect_ant_max5, fh_lect_act_max5, cd_tp_fuente_act_max5, nm_lect_act_max5,");
                        sb.Append("im_ajuste_integr_max5, cd_tp_anoma_max5, fh_lect_ant_max6, cd_tp_fuente_ant_max6, nm_lect_ant_max6, fh_lect_act_max6, cd_tp_fuente_act_max6,");
                        sb.Append("nm_lect_act_max6, im_ajuste_integr_max6, cd_tp_anoma_max6, fh_lect_ant_supervalle, cd_tp_fuente_ant_supervalle, nm_lect_ant_supervalle,");
                        sb.Append("fh_lect_act_supervalle, cd_tp_fuente_act_supervalle, nm_lect_act_supervalle, im_ajuste_integr_supervalle, cd_tp_anoma_supervalle,");
                        sb.Append("cd_concepto_reper_1, im_concepto_reper_1, cd_concepto_reper_2, im_concepto_reper_2, cd_concepto_reper_3, im_concepto_reper_3, cd_concepto_reper_4,");
                        sb.Append("im_concepto_reper_4, cd_concepto_reper_5, im_concepto_reper_5, cd_concepto_reper_6, im_concepto_reper_6, cd_concepto_reper_7, im_concepto_reper_7,");
                        sb.Append("cd_concepto_reper_8, im_concepto_reper_8, cd_concepto_reper_9, im_concepto_reper_9, cd_concepto_reper_10, im_concepto_reper_10, cd_concepto_reper_11,");
                        sb.Append("im_concepto_reper_11, cd_concepto_reper_12, im_concepto_reper_12, cd_cmunicip_atr, de_est_fact, fh_desde_ener, fh_hasta_ener, fh_desde_pot,");
                        sb.Append("fh_hasta_pot, cd_modo_crtl_pot, nm_dias_fact, im_tot_fact, im_saldo_tot_fact, nm_tot_recibos, fh_valor, fh_lim_pago, id_remesa, cd_munic_red,");
                        sb.Append("lg_penaliza_icp, identificador, cd_mensaje, cd_doc_ref, cd_tp_doc, de_tp_doc, de_tp_aceso_serv, lg_relevancia_recl_pdp, de_empr_comer_cne,");
                        sb.Append("cd_huella_fact_atr, fichero, cd_huella_blq_fich, fh_publicacion, fh_public_fichero, cd_solicitud, cd_sec_solicitud, lg_atr, cd_tp_reg, de_tp_reg,");
                        sb.Append("cd_expdte, lg_bloq_af, lg_recl_abta, cd_origen_fact, de_origen_fact, cd_aniofactura, cd_pais, lg_autoconsumo, cd_tp_pm, lg_duracioninfanio,");
                        sb.Append("fh_creacion, cd_usuario_creador, de_usuario_creador, fh_mod, cd_usuario_mod, de_usuario_mod, cd_ubicacion, de_ubicacion, id_exp_af, lg_registro_borrado,");
                        sb.Append("de_tipo_borrado, cd_rg_presion, de_rg_presion, cd_metodo_fact, de_metodo_fact, lg_telemedida, cd_tp_gasinera, de_tp_gasinera, fh_desde_reac, fh_hasta_reac,");
                        sb.Append("de_marca_back, lg_contrato_simul, cd_id_producto, cd_tp_producto, de_tp_producto, lg_arrastre_penali, nm_med_capacitiva, nm_prec_capacitiva, nm_exceso_pot_p1,");
                        sb.Append("nm_exceso_pot_p2, nm_exceso_pot_p3, nm_exceso_pot_p4, nm_exceso_pot_p5, nm_exceso_pot_p6, nm_consumo_medio_5a, nm_consumo_medio) values ");
                        firstOnly = false;
                    }

                    #region Campos
                    if (r["cl_crto_ext_ps"] != System.DBNull.Value)
                        sb.Append("('").Append(r["cl_crto_ext_ps"].ToString()).Append("',");
                    else
                        sb.Append("(null,");

                    if (r["id_crto_ext_ps"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_crto_ext_ps"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_crto_ps"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_crto_ps"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_sec_fact"] != System.DBNull.Value)
                        sb.Append(r["cd_sec_fact"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_linea_neg"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_linea_neg"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_linea_neg"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_linea_neg"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_empr_distr"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_empr_distr"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_est_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_est_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_mes"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_mes"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_fact"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_ini_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_ini_fact"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_fin_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_fin_fact"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_importe"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_importe"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_meses"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_meses"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_fact_sust"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_fact_sust"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_cups"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_cups"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_cups_ext"] != System.DBNull.Value)
                    {
                        sb.Append("'").Append(r["cd_cups_ext"].ToString().Substring(0,20)).Append("',");
                        sb.Append("'").Append(r["cd_cups_ext"].ToString()).Append("',");
                    }                        
                    else
                        sb.Append("null, null,");

                    if (r["cd_empr_distr_cne"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_empr_distr_cne"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_empr_distr_cne"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_empr_distr_cne"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    //nm_med_potencia_
                    for (int i = 0; i < 10; i++)
                    {
                        if (r["nm_med_potencia_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_med_potencia_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["nm_prec_potencia_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_prec_potencia_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    // nm_med_activa_
                    for (int i = 0; i < 10; i++)
                    {
                        if (r["nm_med_activa_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_med_activa_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["nm_prec_activa_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_prec_activa_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    //nm_med_reactiva_
                    for (int i = 0; i < 10; i++)
                    {
                        if (r["nm_med_reactiva_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_med_reactiva_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["nm_prec_reactiva_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_prec_reactiva_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    //cd_concepto_
                    for (int i = 0; i < 20; i++)
                    {
                        if (r["cd_concepto_" + (i + 1)] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_concepto_" + (i + 1)].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_concepto_" + (i + 1) + "_sce"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_concepto_" + (i + 1) + "_sce"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["de_concepto_" + (i + 1)] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_concepto_" + (i + 1)].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["im_concepto_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["im_concepto_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    // de_impuesto_1
                    for (int i = 0; i < 5; i++)
                    {
                        if (r["de_impuesto_" + (i + 1)] != System.DBNull.Value)
                            sb.Append("'").Append(r["de_impuesto_" + (i + 1)].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_porcentaje_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_porcentaje_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["nm_base_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_base_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["nm_importe_" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_importe_" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    if (r["cod_carga_ods"] != System.DBNull.Value)
                        sb.Append(r["cod_carga_ods"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_act_ods"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_act_ods"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_act_dmco"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_act_dmco"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tarifa"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tarifa"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tarifa_ff"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tarifa_ff"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_pot_ctatada"] != System.DBNull.Value)
                        sb.Append(r["nm_pot_ctatada"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_comercializadora"] != System.DBNull.Value)
                        sb.Append(r["cd_comercializadora"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_tipo_consumo"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tipo_consumo"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_municipio"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tipo_consumo"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_cups20_metra"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_cups20_metra"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_teleg"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_teleg"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["consumo_tot_act"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_tot_act"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_punta"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_punta"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_llano"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_llano"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_valle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_valle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_activa4"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_activa4"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_activa5"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_activa5"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_activa6"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_activa6"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_tot_react"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_tot_react"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    for (int i = 0; i < 6; i++)
                    {
                        if (r["consumo_reactiva" + (i + 1)] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["consumo_reactiva" + (i + 1)]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    if (r["potencia_maxpunta"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["potencia_maxpunta"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["potencia_maxllano"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["potencia_maxllano"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["potencia_maxvalle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["potencia_maxvalle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["potencia_max4"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["potencia_max4"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["potencia_max5"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["potencia_max5"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["potencia_max6"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["potencia_max6"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["consumo_supervalle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["consumo_supervalle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_recepcion"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_recepcion"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cod_carga"] != System.DBNull.Value)
                        sb.Append(r["cod_carga"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["id_pseudofactura"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_pseudofactura"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_comercializadora_fact"] != System.DBNull.Value)
                        sb.Append(r["cd_comercializadora_fact"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_empr_comer_cne"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_empr_comer_cne"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_fact_recti"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_fact_recti"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_nif_dist"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_nif_dist"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_comcnmc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_comcnmc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["im_salfact"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_salfact"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["tp_moneda"] != System.DBNull.Value)
                        sb.Append("'").Append(r["tp_moneda"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_feboe"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_feboe"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_med_perdida"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_med_perdida"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_vastrafo"] != System.DBNull.Value)
                        sb.Append(r["nm_vastrafo"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_porc_perdida"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_porc_perdida"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_fact_curva"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_fact_curva"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_desde_cch"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_desde_cch"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_hasta_cch"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_hasta_cch"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_num_mes"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_num_mes"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_ant_punta"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_punta"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_punta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_punta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_punta"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_punta"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_act_punta"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_punta"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_punta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_punta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_punta"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_punta"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_punta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_punta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_punta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_punta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_ant_llano"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_llano"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_llano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_llano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_llano"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_llano"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_act_llano"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_llano"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_llano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_llano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_llano"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_llano"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_llano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_llano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_llano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_llano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_ant_valle"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_valle"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_valle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_valle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_valle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_valle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_act_valle"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_valle"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_valle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_valle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_valle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_valle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_valle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_valle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_valle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_valle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    for (int i = 4; i <= 6; i++)
                    {
                        if (r["fh_lect_ant_activa" + i] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_activa" + i]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_fuente_ant_activa" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_fuente_ant_activa" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_lect_ant_activa" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_lect_ant_activa" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["fh_lect_act_activa" + i] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_activa" + i]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_fuente_act_activa" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_fuente_act_activa" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_lect_act_activa" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_lect_act_activa" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["im_ajuste_integr_activa" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["im_ajuste_integr_activa" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_anoma_activa" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_anoma_activa" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");
                    }

                    // fh_lect_ant_reactiva
                    for (int i = 1; i <= 6; i++)
                    {
                        if (r["fh_lect_ant_reactiva" + i] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_reactiva" + i]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_fuente_ant_reactiva" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_fuente_ant_reactiva" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_lect_ant_reactiva" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_lect_ant_reactiva" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["fh_lect_act_reactiva" + i] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_reactiva" + i]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_fuente_act_reactiva" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_fuente_act_reactiva" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_lect_act_reactiva" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_lect_act_reactiva" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["im_ajuste_integr_reactiva" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["im_ajuste_integr_reactiva" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_anoma_reactiva" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_anoma_reactiva" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");
                    }


                    if (r["fh_lect_ant_maxpunta"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_maxpunta"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_maxpunta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_maxpunta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_maxpunta"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_maxpunta"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_act_maxpunta"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_maxpunta"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_maxpunta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_maxpunta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_maxpunta"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_maxpunta"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_maxpunta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_maxpunta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_maxpunta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_maxpunta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_ant_maxllano"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_maxllano"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_maxllano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_maxllano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_maxllano"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_maxllano"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");


                    if (r["fh_lect_act_maxllano"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_maxllano"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_maxllano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_maxllano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_maxllano"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_maxllano"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_maxllano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_maxllano"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_maxllano"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_maxllano"].ToString()).Append("',");
                    else
                        sb.Append("null,");


                    if (r["fh_lect_ant_maxvalle"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_maxvalle"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_maxvalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_maxvalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_maxvalle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_maxvalle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_act_maxvalle"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_maxvalle"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_maxvalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_maxvalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_maxvalle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_maxvalle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_maxvalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_maxvalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_maxvalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_maxvalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    // fh_lect_ant_max4
                    for (int i = 4; i <= 6; i++)
                    {
                        if (r["fh_lect_ant_max" + i] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_max" + i]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_fuente_ant_max" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_fuente_ant_max" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_lect_ant_max" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_lect_ant_max" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["fh_lect_act_max" + i] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_max" + i]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_fuente_act_max" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_fuente_act_max" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["nm_lect_act_max" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_lect_act_max" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["im_ajuste_integr_max" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["im_ajuste_integr_max" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_anoma_max" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_anoma_max" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");
                    }

                    if (r["fh_lect_ant_supervalle"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_ant_supervalle"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_ant_supervalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_ant_supervalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_ant_supervalle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_ant_supervalle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lect_act_supervalle"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lect_act_supervalle"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_act_supervalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_act_supervalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_lect_act_supervalle"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_lect_act_supervalle"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_ajuste_integr_supervalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["im_ajuste_integr_supervalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_anoma_supervalle"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_anoma_supervalle"].ToString()).Append("',");
                    else
                        sb.Append("null,");


                    // cd_concepto_reper_
                    for (int i = 1; i <= 12; i++)
                    {
                        if (r["cd_concepto_reper_" + i] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_concepto_reper_" + i].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["im_concepto_reper_" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["im_concepto_reper_" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }


                    if (r["cd_cmunicip_atr"] != System.DBNull.Value)
                        sb.Append(r["cd_cmunicip_atr"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["de_est_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_est_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_desde_ener"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_desde_ener"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_hasta_ener"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_hasta_ener"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_desde_pot"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_desde_pot"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_hasta_pot"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_hasta_pot"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");


                    if (r["cd_modo_crtl_pot"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_modo_crtl_pot"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_dias_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["nm_dias_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["im_tot_fact"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_tot_fact"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_saldo_tot_fact"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_saldo_tot_fact"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_tot_recibos"] != System.DBNull.Value)
                        sb.Append(r["nm_tot_recibos"].ToString()).Append(",");
                    else
                        sb.Append("null,");


                    if (r["fh_valor"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_valor"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_lim_pago"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lim_pago"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");


                    if (r["id_remesa"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_remesa"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_munic_red"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_munic_red"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_penaliza_icp"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_penaliza_icp"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["identificador"] != System.DBNull.Value)
                        sb.Append("'").Append(r["identificador"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_mensaje"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_mensaje"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_doc_ref"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_doc_ref"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_doc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_doc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_doc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_doc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_aceso_serv"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_aceso_serv"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_relevancia_recl_pdp"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_relevancia_recl_pdp"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_empr_comer_cne"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_empr_comer_cne"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_huella_fact_atr"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_huella_fact_atr"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fichero"] != System.DBNull.Value)
                        sb.Append("'").Append(r["fichero"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_huella_blq_fich"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_huella_blq_fich"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_publicacion"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_publicacion"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_public_fichero"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_public_fichero"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_solicitud"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_solicitud"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_sec_solicitud"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_sec_solicitud"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_atr"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_atr"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_reg"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_reg"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_reg"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_reg"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_expdte"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_expdte"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_bloq_af"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_bloq_af"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_recl_abta"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_recl_abta"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_origen_fact"] != System.DBNull.Value)
                        sb.Append(r["cd_origen_fact"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["de_origen_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_origen_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_aniofactura"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_aniofactura"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_pais"] != System.DBNull.Value)
                        sb.Append(r["cd_pais"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["lg_autoconsumo"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_autoconsumo"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_pm"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_pm"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_duracioninfanio"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_duracioninfanio"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_creacion"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_creacion"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_usuario_creador"] != System.DBNull.Value)
                        sb.Append(r["cd_usuario_creador"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["de_usuario_creador"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_usuario_creador"].ToString()).Append("',");
                    else
                        sb.Append("null,");


                    if (r["fh_mod"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_mod"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_usuario_mod"] != System.DBNull.Value)
                        sb.Append(r["cd_usuario_mod"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["de_usuario_mod"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_usuario_mod"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_ubicacion"] != System.DBNull.Value)
                        sb.Append(r["cd_ubicacion"].ToString()).Append(",");
                    else
                        sb.Append("null,");


                    if (r["de_ubicacion"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_ubicacion"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_exp_af"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_exp_af"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_registro_borrado"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_registro_borrado"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tipo_borrado"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tipo_borrado"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_rg_presion"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_rg_presion"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_rg_presion"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_rg_presion"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_metodo_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_metodo_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_metodo_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_metodo_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_telemedida"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_telemedida"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_gasinera"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_gasinera"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_gasinera"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_gasinera"].ToString()).Append("',");
                    else
                        sb.Append("null,");


                    if (r["fh_desde_reac"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_desde_reac"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_hasta_reac"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_hasta_reac"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_marca_back"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_marca_back"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_contrato_simul"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_contrato_simul"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_id_producto"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_id_producto"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_producto"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_producto"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_producto"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_producto"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_arrastre_penali"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_arrastre_penali"].ToString()).Append("',");
                    else
                        sb.Append("null,");


                    if (r["nm_med_capacitiva"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_med_capacitiva"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_prec_capacitiva"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_prec_capacitiva"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");



                    for (int i = 1; i <= 6; i++)
                    {
                        if (r["nm_exceso_pot_p" + i] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["nm_exceso_pot_p" + i]).ToString("").Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }

                    if (r["nm_consumo_medio_5a"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_consumo_medio_5a"]).ToString("").Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_consumo_medio"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_consumo_medio"]).ToString("").Replace(",", ".")).Append("),");
                    else
                        sb.Append("null),");


                    #endregion


                    if (j == 100)
                    {
                        Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                        firstOnly = true;
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
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
                    dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                    commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                    commandmy.ExecuteNonQuery();
                    dbmy.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    j = 0;
                }

                // Construimos conceptos
                for (int i = 1; i < 21; i++)
                {
                    strSql = "REPLACE INTO t_ed_h_facts_atr_ml_conceptos"
                        + " SELECT f.id_fact, f.cd_cups_ext,"
                        + " f.cd_concepto_" + i + ","
                        + " f.cd_concepto_" + i + "_sce,"
                        + " f.de_concepto_" + i + ","
                        + " f.im_concepto_" + i
                        + " FROM t_ed_h_facts_atr_ml_tmp f"
                        + " where f.cd_concepto_" + i + " <> ''";
                    Console.WriteLine("Ejecutando " + strSql);
                    ficheroLog.Add("Ejecutando " + strSql);
                    dbmy = new MySQLDB(MySQLDB.Esquemas.MED);
                    commandmy = new MySqlCommand(strSql, dbmy.con);
                    commandmy.ExecuteNonQuery();
                    dbmy.CloseConnection();
                }

                // Volcamos tmp sobre tabla normal

                strSql = "replace into t_ed_h_facts_atr_ml select * from t_ed_h_facts_atr_ml_tmp";
                Console.WriteLine("Ejecutando " + strSql);
                ficheroLog.Add("Ejecutando " + strSql);
                dbmy = new MySQLDB(MySQLDB.Esquemas.MED);
                commandmy = new MySqlCommand(strSql, dbmy.con);
                commandmy.ExecuteNonQuery();
                dbmy.CloseConnection();


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                ficheroLog.addError(ex.Message);
            }
        }

        private string Consulta_Proceso_Inicial(DateTime fd, string cups13)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT cl_crto_ext_ps, id_crto_ext_ps, id_crto_ps, cd_sec_fact, cd_linea_neg, de_linea_neg, cd_empr_distr, cd_est_fact, cd_mes, id_fact, fh_fact, fh_ini_fact, fh_fin_fact,");
            sb.Append("nm_importe, nm_meses, cd_tp_fact, de_tp_fact, id_fact_sust, cd_cups, cd_cups_ext, cd_empr_distr_cne, de_empr_distr_cne, nm_med_potencia_1, nm_prec_potencia_1, nm_med_potencia_2, nm_prec_potencia_2,");
            sb.Append("nm_med_potencia_3, nm_prec_potencia_3, nm_med_potencia_4, nm_prec_potencia_4, nm_med_potencia_5, nm_prec_potencia_5, nm_med_potencia_6, nm_prec_potencia_6, nm_med_potencia_7, nm_prec_potencia_7,");
            sb.Append("nm_med_potencia_8, nm_prec_potencia_8, nm_med_potencia_9, nm_prec_potencia_9, nm_med_potencia_10, nm_prec_potencia_10, nm_med_activa_1, nm_prec_activa_1, nm_med_activa_2, nm_prec_activa_2,");
            sb.Append("nm_med_activa_3, nm_prec_activa_3, nm_med_activa_4, nm_prec_activa_4, nm_med_activa_5, nm_prec_activa_5, nm_med_activa_6, nm_prec_activa_6, nm_med_activa_7, nm_prec_activa_7, nm_med_activa_8,");
            sb.Append("nm_prec_activa_8, nm_med_activa_9, nm_prec_activa_9, nm_med_activa_10, nm_prec_activa_10, nm_med_reactiva_1, nm_prec_reactiva_1, nm_med_reactiva_2, nm_prec_reactiva_2, nm_med_reactiva_3,");
            sb.Append("nm_prec_reactiva_3, nm_med_reactiva_4, nm_prec_reactiva_4, nm_med_reactiva_5, nm_prec_reactiva_5, nm_med_reactiva_6, nm_prec_reactiva_6, nm_med_reactiva_7, nm_prec_reactiva_7, nm_med_reactiva_8,");
            sb.Append("nm_prec_reactiva_8, nm_med_reactiva_9, nm_prec_reactiva_9, nm_med_reactiva_10, nm_prec_reactiva_10, cd_concepto_1, cd_concepto_1_sce, de_concepto_1, im_concepto_1, cd_concepto_2, cd_concepto_2_sce,");
            sb.Append("de_concepto_2, im_concepto_2, cd_concepto_3, cd_concepto_3_sce, de_concepto_3, im_concepto_3, cd_concepto_4, cd_concepto_4_sce, de_concepto_4, im_concepto_4, cd_concepto_5, cd_concepto_5_sce, de_concepto_5,");
            sb.Append("im_concepto_5, cd_concepto_6, cd_concepto_6_sce, de_concepto_6, im_concepto_6, cd_concepto_7, cd_concepto_7_sce, de_concepto_7, im_concepto_7, cd_concepto_8, cd_concepto_8_sce, de_concepto_8, im_concepto_8,");
            sb.Append("cd_concepto_9, cd_concepto_9_sce, de_concepto_9, im_concepto_9, cd_concepto_10, cd_concepto_10_sce, de_concepto_10, im_concepto_10, cd_concepto_11, cd_concepto_11_sce, de_concepto_11, im_concepto_11,");
            sb.Append("cd_concepto_12, cd_concepto_12_sce, de_concepto_12, im_concepto_12, cd_concepto_13, cd_concepto_13_sce, de_concepto_13, im_concepto_13, cd_concepto_14, cd_concepto_14_sce, de_concepto_14, im_concepto_14,");
            sb.Append("cd_concepto_15, cd_concepto_15_sce, de_concepto_15, im_concepto_15, cd_concepto_16, cd_concepto_16_sce, de_concepto_16, im_concepto_16, cd_concepto_17, cd_concepto_17_sce, de_concepto_17, im_concepto_17,");
            sb.Append("cd_concepto_18, cd_concepto_18_sce, de_concepto_18, im_concepto_18, cd_concepto_19, cd_concepto_19_sce, de_concepto_19, im_concepto_19, cd_concepto_20, cd_concepto_20_sce, de_concepto_20, im_concepto_20,");
            sb.Append("de_impuesto_1, nm_porcentaje_1, nm_base_1, nm_importe_1, de_impuesto_2, nm_porcentaje_2, nm_base_2, nm_importe_2, de_impuesto_3, nm_porcentaje_3, nm_base_3, nm_importe_3, de_impuesto_4, nm_porcentaje_4,");
            sb.Append("nm_base_4, nm_importe_4, de_impuesto_5, nm_porcentaje_5, nm_base_5, nm_importe_5, cod_carga_ods, fh_act_ods, fh_act_dmco, cd_tarifa, cd_tarifa_ff, nm_pot_ctatada, cd_comercializadora, cd_tipo_consumo,");
            sb.Append("cd_municipio, cd_cups20_metra, cd_tp_teleg, consumo_tot_act, consumo_punta, consumo_llano, consumo_valle, consumo_activa4, consumo_activa5, consumo_activa6, consumo_tot_react, consumo_reactiva1, consumo_reactiva2,");
            sb.Append("consumo_reactiva3, consumo_reactiva4, consumo_reactiva5, consumo_reactiva6, potencia_maxpunta, potencia_maxllano, potencia_maxvalle, potencia_max4, potencia_max5, potencia_max6, consumo_supervalle, fh_recepcion,");
            sb.Append("cod_carga, id_pseudofactura, cd_comercializadora_fact, cd_empr_comer_cne, id_fact_recti, id_nif_dist, de_comcnmc, im_salfact, tp_moneda, fh_feboe, lg_med_perdida, nm_vastrafo, nm_porc_perdida, cd_fact_curva,");
            sb.Append("fh_desde_cch, fh_hasta_cch, nm_num_mes, fh_lect_ant_punta, cd_tp_fuente_ant_punta, nm_lect_ant_punta, fh_lect_act_punta, cd_tp_fuente_act_punta, nm_lect_act_punta, im_ajuste_integr_punta, cd_tp_anoma_punta,");
            sb.Append("fh_lect_ant_llano, cd_tp_fuente_ant_llano, nm_lect_ant_llano, fh_lect_act_llano, cd_tp_fuente_act_llano, nm_lect_act_llano, im_ajuste_integr_llano, cd_tp_anoma_llano, fh_lect_ant_valle, cd_tp_fuente_ant_valle,");
            sb.Append("nm_lect_ant_valle, fh_lect_act_valle, cd_tp_fuente_act_valle, nm_lect_act_valle, im_ajuste_integr_valle, cd_tp_anoma_valle, fh_lect_ant_activa4, cd_tp_fuente_ant_activa4, nm_lect_ant_activa4, fh_lect_act_activa4,");
            sb.Append("cd_tp_fuente_act_activa4, nm_lect_act_activa4, im_ajuste_integr_activa4, cd_tp_anoma_activa4, fh_lect_ant_activa5, cd_tp_fuente_ant_activa5, nm_lect_ant_activa5, fh_lect_act_activa5, cd_tp_fuente_act_activa5,");
            sb.Append("nm_lect_act_activa5, im_ajuste_integr_activa5, cd_tp_anoma_activa5, fh_lect_ant_activa6, cd_tp_fuente_ant_activa6, nm_lect_ant_activa6, fh_lect_act_activa6, cd_tp_fuente_act_activa6, nm_lect_act_activa6,");
            sb.Append("im_ajuste_integr_activa6, cd_tp_anoma_activa6, fh_lect_ant_reactiva1, cd_tp_fuente_ant_reactiva1, nm_lect_ant_reactiva1, fh_lect_act_reactiva1, cd_tp_fuente_act_reactiva1, nm_lect_act_reactiva1,");
            sb.Append("im_ajuste_integr_reactiva1, cd_tp_anoma_reactiva1, fh_lect_ant_reactiva2, cd_tp_fuente_ant_reactiva2, nm_lect_ant_reactiva2, fh_lect_act_reactiva2, cd_tp_fuente_act_reactiva2, nm_lect_act_reactiva2,");
            sb.Append("im_ajuste_integr_reactiva2, cd_tp_anoma_reactiva2, fh_lect_ant_reactiva3, cd_tp_fuente_ant_reactiva3, nm_lect_ant_reactiva3, fh_lect_act_reactiva3, cd_tp_fuente_act_reactiva3, nm_lect_act_reactiva3,");
            sb.Append("im_ajuste_integr_reactiva3, cd_tp_anoma_reactiva3, fh_lect_ant_reactiva4, cd_tp_fuente_ant_reactiva4, nm_lect_ant_reactiva4, fh_lect_act_reactiva4, cd_tp_fuente_act_reactiva4, nm_lect_act_reactiva4,");
            sb.Append("im_ajuste_integr_reactiva4, cd_tp_anoma_reactiva4, fh_lect_ant_reactiva5, cd_tp_fuente_ant_reactiva5, nm_lect_ant_reactiva5, fh_lect_act_reactiva5, cd_tp_fuente_act_reactiva5, nm_lect_act_reactiva5,");
            sb.Append("im_ajuste_integr_reactiva5, cd_tp_anoma_reactiva5, fh_lect_ant_reactiva6, cd_tp_fuente_ant_reactiva6, nm_lect_ant_reactiva6, fh_lect_act_reactiva6, cd_tp_fuente_act_reactiva6, nm_lect_act_reactiva6,");
            sb.Append("im_ajuste_integr_reactiva6, cd_tp_anoma_reactiva6, fh_lect_ant_maxpunta, cd_tp_fuente_ant_maxpunta, nm_lect_ant_maxpunta, fh_lect_act_maxpunta, cd_tp_fuente_act_maxpunta, nm_lect_act_maxpunta,");
            sb.Append("im_ajuste_integr_maxpunta, cd_tp_anoma_maxpunta, fh_lect_ant_maxllano, cd_tp_fuente_ant_maxllano, nm_lect_ant_maxllano, fh_lect_act_maxllano, cd_tp_fuente_act_maxllano, nm_lect_act_maxllano,");
            sb.Append("im_ajuste_integr_maxllano, cd_tp_anoma_maxllano, fh_lect_ant_maxvalle, cd_tp_fuente_ant_maxvalle, nm_lect_ant_maxvalle, fh_lect_act_maxvalle, cd_tp_fuente_act_maxvalle, nm_lect_act_maxvalle,");
            sb.Append("im_ajuste_integr_maxvalle, cd_tp_anoma_maxvalle, fh_lect_ant_max4, cd_tp_fuente_ant_max4, nm_lect_ant_max4, fh_lect_act_max4, cd_tp_fuente_act_max4, nm_lect_act_max4, im_ajuste_integr_max4,");
            sb.Append("cd_tp_anoma_max4, fh_lect_ant_max5, cd_tp_fuente_ant_max5, nm_lect_ant_max5, fh_lect_act_max5, cd_tp_fuente_act_max5, nm_lect_act_max5, im_ajuste_integr_max5, cd_tp_anoma_max5, fh_lect_ant_max6,");
            sb.Append("cd_tp_fuente_ant_max6, nm_lect_ant_max6, fh_lect_act_max6, cd_tp_fuente_act_max6, nm_lect_act_max6, im_ajuste_integr_max6, cd_tp_anoma_max6, fh_lect_ant_supervalle, cd_tp_fuente_ant_supervalle,");
            sb.Append("nm_lect_ant_supervalle, fh_lect_act_supervalle, cd_tp_fuente_act_supervalle, nm_lect_act_supervalle, im_ajuste_integr_supervalle, cd_tp_anoma_supervalle, cd_concepto_reper_1, im_concepto_reper_1,");
            sb.Append("cd_concepto_reper_2, im_concepto_reper_2, cd_concepto_reper_3, im_concepto_reper_3, cd_concepto_reper_4, im_concepto_reper_4, cd_concepto_reper_5, im_concepto_reper_5, cd_concepto_reper_6, im_concepto_reper_6,");
            sb.Append("cd_concepto_reper_7, im_concepto_reper_7, cd_concepto_reper_8, im_concepto_reper_8, cd_concepto_reper_9, im_concepto_reper_9, cd_concepto_reper_10, im_concepto_reper_10, cd_concepto_reper_11, im_concepto_reper_11,");
            sb.Append("cd_concepto_reper_12, im_concepto_reper_12, cd_cmunicip_atr, de_est_fact, fh_desde_ener, fh_hasta_ener, fh_desde_pot, fh_hasta_pot, cd_modo_crtl_pot, nm_dias_fact, im_tot_fact, im_saldo_tot_fact, nm_tot_recibos,");
            sb.Append("fh_valor, fh_lim_pago, id_remesa, cd_munic_red, lg_penaliza_icp, identificador, cd_mensaje, cd_doc_ref, cd_tp_doc, de_tp_doc, de_tp_aceso_serv, lg_relevancia_recl_pdp, de_empr_comer_cne, cd_huella_fact_atr, fichero,");
            sb.Append("cd_huella_blq_fich, fh_publicacion, fh_public_fichero, cd_solicitud, cd_sec_solicitud, lg_atr, cd_tp_reg, de_tp_reg, cd_expdte, lg_bloq_af, lg_recl_abta, cd_origen_fact, de_origen_fact, cd_aniofactura, cd_pais,");
            sb.Append("lg_autoconsumo, cd_tp_pm, lg_duracioninfanio, fh_creacion, cd_usuario_creador, de_usuario_creador, fh_mod, cd_usuario_mod, de_usuario_mod, cd_ubicacion, de_ubicacion, id_exp_af, lg_registro_borrado, de_tipo_borrado,");
            sb.Append("cd_rg_presion, de_rg_presion, cd_metodo_fact, de_metodo_fact, lg_telemedida, cd_tp_gasinera, de_tp_gasinera, fh_desde_reac, fh_hasta_reac, de_marca_back, lg_contrato_simul, cd_id_producto, cd_tp_producto,");
            sb.Append("de_tp_producto, lg_arrastre_penali, nm_med_capacitiva, nm_prec_capacitiva, nm_exceso_pot_p1, nm_exceso_pot_p2, nm_exceso_pot_p3, nm_exceso_pot_p4, nm_exceso_pot_p5, nm_exceso_pot_p6, nm_consumo_medio_5a,");
            sb.Append("nm_consumo_medio");
            sb.Append(" from ed_owner.t_ed_h_facts_atr_ml ml where");
            sb.Append(" cd_cups = '" + cups13 + "' AND");
            sb.Append(" ml.fh_act_dmco >= '").Append(fd.ToString("yyyy-MM-dd")).Append("'");

            return sb.ToString();
        }

        private string Consulta_Peajes(DateTime fd, List<string> lista_cups13)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT cl_crto_ext_ps, id_crto_ext_ps, id_crto_ps, cd_sec_fact, cd_linea_neg, de_linea_neg, cd_empr_distr, cd_est_fact, cd_mes, id_fact, fh_fact, fh_ini_fact, fh_fin_fact,");
            sb.Append("nm_importe, nm_meses, cd_tp_fact, de_tp_fact, id_fact_sust, cd_cups, cd_cups_ext, cd_empr_distr_cne, de_empr_distr_cne, nm_med_potencia_1, nm_prec_potencia_1, nm_med_potencia_2, nm_prec_potencia_2,");
            sb.Append("nm_med_potencia_3, nm_prec_potencia_3, nm_med_potencia_4, nm_prec_potencia_4, nm_med_potencia_5, nm_prec_potencia_5, nm_med_potencia_6, nm_prec_potencia_6, nm_med_potencia_7, nm_prec_potencia_7,");
            sb.Append("nm_med_potencia_8, nm_prec_potencia_8, nm_med_potencia_9, nm_prec_potencia_9, nm_med_potencia_10, nm_prec_potencia_10, nm_med_activa_1, nm_prec_activa_1, nm_med_activa_2, nm_prec_activa_2,");
            sb.Append("nm_med_activa_3, nm_prec_activa_3, nm_med_activa_4, nm_prec_activa_4, nm_med_activa_5, nm_prec_activa_5, nm_med_activa_6, nm_prec_activa_6, nm_med_activa_7, nm_prec_activa_7, nm_med_activa_8,");
            sb.Append("nm_prec_activa_8, nm_med_activa_9, nm_prec_activa_9, nm_med_activa_10, nm_prec_activa_10, nm_med_reactiva_1, nm_prec_reactiva_1, nm_med_reactiva_2, nm_prec_reactiva_2, nm_med_reactiva_3,");
            sb.Append("nm_prec_reactiva_3, nm_med_reactiva_4, nm_prec_reactiva_4, nm_med_reactiva_5, nm_prec_reactiva_5, nm_med_reactiva_6, nm_prec_reactiva_6, nm_med_reactiva_7, nm_prec_reactiva_7, nm_med_reactiva_8,");
            sb.Append("nm_prec_reactiva_8, nm_med_reactiva_9, nm_prec_reactiva_9, nm_med_reactiva_10, nm_prec_reactiva_10, cd_concepto_1, cd_concepto_1_sce, de_concepto_1, im_concepto_1, cd_concepto_2, cd_concepto_2_sce,");
            sb.Append("de_concepto_2, im_concepto_2, cd_concepto_3, cd_concepto_3_sce, de_concepto_3, im_concepto_3, cd_concepto_4, cd_concepto_4_sce, de_concepto_4, im_concepto_4, cd_concepto_5, cd_concepto_5_sce, de_concepto_5,");
            sb.Append("im_concepto_5, cd_concepto_6, cd_concepto_6_sce, de_concepto_6, im_concepto_6, cd_concepto_7, cd_concepto_7_sce, de_concepto_7, im_concepto_7, cd_concepto_8, cd_concepto_8_sce, de_concepto_8, im_concepto_8,");
            sb.Append("cd_concepto_9, cd_concepto_9_sce, de_concepto_9, im_concepto_9, cd_concepto_10, cd_concepto_10_sce, de_concepto_10, im_concepto_10, cd_concepto_11, cd_concepto_11_sce, de_concepto_11, im_concepto_11,");
            sb.Append("cd_concepto_12, cd_concepto_12_sce, de_concepto_12, im_concepto_12, cd_concepto_13, cd_concepto_13_sce, de_concepto_13, im_concepto_13, cd_concepto_14, cd_concepto_14_sce, de_concepto_14, im_concepto_14,");
            sb.Append("cd_concepto_15, cd_concepto_15_sce, de_concepto_15, im_concepto_15, cd_concepto_16, cd_concepto_16_sce, de_concepto_16, im_concepto_16, cd_concepto_17, cd_concepto_17_sce, de_concepto_17, im_concepto_17,");
            sb.Append("cd_concepto_18, cd_concepto_18_sce, de_concepto_18, im_concepto_18, cd_concepto_19, cd_concepto_19_sce, de_concepto_19, im_concepto_19, cd_concepto_20, cd_concepto_20_sce, de_concepto_20, im_concepto_20,");
            sb.Append("de_impuesto_1, nm_porcentaje_1, nm_base_1, nm_importe_1, de_impuesto_2, nm_porcentaje_2, nm_base_2, nm_importe_2, de_impuesto_3, nm_porcentaje_3, nm_base_3, nm_importe_3, de_impuesto_4, nm_porcentaje_4,");
            sb.Append("nm_base_4, nm_importe_4, de_impuesto_5, nm_porcentaje_5, nm_base_5, nm_importe_5, cod_carga_ods, fh_act_ods, fh_act_dmco, cd_tarifa, cd_tarifa_ff, nm_pot_ctatada, cd_comercializadora, cd_tipo_consumo,");
            sb.Append("cd_municipio, cd_cups20_metra, cd_tp_teleg, consumo_tot_act, consumo_punta, consumo_llano, consumo_valle, consumo_activa4, consumo_activa5, consumo_activa6, consumo_tot_react, consumo_reactiva1, consumo_reactiva2,");
            sb.Append("consumo_reactiva3, consumo_reactiva4, consumo_reactiva5, consumo_reactiva6, potencia_maxpunta, potencia_maxllano, potencia_maxvalle, potencia_max4, potencia_max5, potencia_max6, consumo_supervalle, fh_recepcion,");
            sb.Append("cod_carga, id_pseudofactura, cd_comercializadora_fact, cd_empr_comer_cne, id_fact_recti, id_nif_dist, de_comcnmc, im_salfact, tp_moneda, fh_feboe, lg_med_perdida, nm_vastrafo, nm_porc_perdida, cd_fact_curva,");
            sb.Append("fh_desde_cch, fh_hasta_cch, nm_num_mes, fh_lect_ant_punta, cd_tp_fuente_ant_punta, nm_lect_ant_punta, fh_lect_act_punta, cd_tp_fuente_act_punta, nm_lect_act_punta, im_ajuste_integr_punta, cd_tp_anoma_punta,");
            sb.Append("fh_lect_ant_llano, cd_tp_fuente_ant_llano, nm_lect_ant_llano, fh_lect_act_llano, cd_tp_fuente_act_llano, nm_lect_act_llano, im_ajuste_integr_llano, cd_tp_anoma_llano, fh_lect_ant_valle, cd_tp_fuente_ant_valle,");
            sb.Append("nm_lect_ant_valle, fh_lect_act_valle, cd_tp_fuente_act_valle, nm_lect_act_valle, im_ajuste_integr_valle, cd_tp_anoma_valle, fh_lect_ant_activa4, cd_tp_fuente_ant_activa4, nm_lect_ant_activa4, fh_lect_act_activa4,");
            sb.Append("cd_tp_fuente_act_activa4, nm_lect_act_activa4, im_ajuste_integr_activa4, cd_tp_anoma_activa4, fh_lect_ant_activa5, cd_tp_fuente_ant_activa5, nm_lect_ant_activa5, fh_lect_act_activa5, cd_tp_fuente_act_activa5,");
            sb.Append("nm_lect_act_activa5, im_ajuste_integr_activa5, cd_tp_anoma_activa5, fh_lect_ant_activa6, cd_tp_fuente_ant_activa6, nm_lect_ant_activa6, fh_lect_act_activa6, cd_tp_fuente_act_activa6, nm_lect_act_activa6,");
            sb.Append("im_ajuste_integr_activa6, cd_tp_anoma_activa6, fh_lect_ant_reactiva1, cd_tp_fuente_ant_reactiva1, nm_lect_ant_reactiva1, fh_lect_act_reactiva1, cd_tp_fuente_act_reactiva1, nm_lect_act_reactiva1,");
            sb.Append("im_ajuste_integr_reactiva1, cd_tp_anoma_reactiva1, fh_lect_ant_reactiva2, cd_tp_fuente_ant_reactiva2, nm_lect_ant_reactiva2, fh_lect_act_reactiva2, cd_tp_fuente_act_reactiva2, nm_lect_act_reactiva2,");
            sb.Append("im_ajuste_integr_reactiva2, cd_tp_anoma_reactiva2, fh_lect_ant_reactiva3, cd_tp_fuente_ant_reactiva3, nm_lect_ant_reactiva3, fh_lect_act_reactiva3, cd_tp_fuente_act_reactiva3, nm_lect_act_reactiva3,");
            sb.Append("im_ajuste_integr_reactiva3, cd_tp_anoma_reactiva3, fh_lect_ant_reactiva4, cd_tp_fuente_ant_reactiva4, nm_lect_ant_reactiva4, fh_lect_act_reactiva4, cd_tp_fuente_act_reactiva4, nm_lect_act_reactiva4,");
            sb.Append("im_ajuste_integr_reactiva4, cd_tp_anoma_reactiva4, fh_lect_ant_reactiva5, cd_tp_fuente_ant_reactiva5, nm_lect_ant_reactiva5, fh_lect_act_reactiva5, cd_tp_fuente_act_reactiva5, nm_lect_act_reactiva5,");
            sb.Append("im_ajuste_integr_reactiva5, cd_tp_anoma_reactiva5, fh_lect_ant_reactiva6, cd_tp_fuente_ant_reactiva6, nm_lect_ant_reactiva6, fh_lect_act_reactiva6, cd_tp_fuente_act_reactiva6, nm_lect_act_reactiva6,");
            sb.Append("im_ajuste_integr_reactiva6, cd_tp_anoma_reactiva6, fh_lect_ant_maxpunta, cd_tp_fuente_ant_maxpunta, nm_lect_ant_maxpunta, fh_lect_act_maxpunta, cd_tp_fuente_act_maxpunta, nm_lect_act_maxpunta,");
            sb.Append("im_ajuste_integr_maxpunta, cd_tp_anoma_maxpunta, fh_lect_ant_maxllano, cd_tp_fuente_ant_maxllano, nm_lect_ant_maxllano, fh_lect_act_maxllano, cd_tp_fuente_act_maxllano, nm_lect_act_maxllano,");
            sb.Append("im_ajuste_integr_maxllano, cd_tp_anoma_maxllano, fh_lect_ant_maxvalle, cd_tp_fuente_ant_maxvalle, nm_lect_ant_maxvalle, fh_lect_act_maxvalle, cd_tp_fuente_act_maxvalle, nm_lect_act_maxvalle,");
            sb.Append("im_ajuste_integr_maxvalle, cd_tp_anoma_maxvalle, fh_lect_ant_max4, cd_tp_fuente_ant_max4, nm_lect_ant_max4, fh_lect_act_max4, cd_tp_fuente_act_max4, nm_lect_act_max4, im_ajuste_integr_max4,");
            sb.Append("cd_tp_anoma_max4, fh_lect_ant_max5, cd_tp_fuente_ant_max5, nm_lect_ant_max5, fh_lect_act_max5, cd_tp_fuente_act_max5, nm_lect_act_max5, im_ajuste_integr_max5, cd_tp_anoma_max5, fh_lect_ant_max6,");
            sb.Append("cd_tp_fuente_ant_max6, nm_lect_ant_max6, fh_lect_act_max6, cd_tp_fuente_act_max6, nm_lect_act_max6, im_ajuste_integr_max6, cd_tp_anoma_max6, fh_lect_ant_supervalle, cd_tp_fuente_ant_supervalle,");
            sb.Append("nm_lect_ant_supervalle, fh_lect_act_supervalle, cd_tp_fuente_act_supervalle, nm_lect_act_supervalle, im_ajuste_integr_supervalle, cd_tp_anoma_supervalle, cd_concepto_reper_1, im_concepto_reper_1,");
            sb.Append("cd_concepto_reper_2, im_concepto_reper_2, cd_concepto_reper_3, im_concepto_reper_3, cd_concepto_reper_4, im_concepto_reper_4, cd_concepto_reper_5, im_concepto_reper_5, cd_concepto_reper_6, im_concepto_reper_6,");
            sb.Append("cd_concepto_reper_7, im_concepto_reper_7, cd_concepto_reper_8, im_concepto_reper_8, cd_concepto_reper_9, im_concepto_reper_9, cd_concepto_reper_10, im_concepto_reper_10, cd_concepto_reper_11, im_concepto_reper_11,");
            sb.Append("cd_concepto_reper_12, im_concepto_reper_12, cd_cmunicip_atr, de_est_fact, fh_desde_ener, fh_hasta_ener, fh_desde_pot, fh_hasta_pot, cd_modo_crtl_pot, nm_dias_fact, im_tot_fact, im_saldo_tot_fact, nm_tot_recibos,");
            sb.Append("fh_valor, fh_lim_pago, id_remesa, cd_munic_red, lg_penaliza_icp, identificador, cd_mensaje, cd_doc_ref, cd_tp_doc, de_tp_doc, de_tp_aceso_serv, lg_relevancia_recl_pdp, de_empr_comer_cne, cd_huella_fact_atr, fichero,");
            sb.Append("cd_huella_blq_fich, fh_publicacion, fh_public_fichero, cd_solicitud, cd_sec_solicitud, lg_atr, cd_tp_reg, de_tp_reg, cd_expdte, lg_bloq_af, lg_recl_abta, cd_origen_fact, de_origen_fact, cd_aniofactura, cd_pais,");
            sb.Append("lg_autoconsumo, cd_tp_pm, lg_duracioninfanio, fh_creacion, cd_usuario_creador, de_usuario_creador, fh_mod, cd_usuario_mod, de_usuario_mod, cd_ubicacion, de_ubicacion, id_exp_af, lg_registro_borrado, de_tipo_borrado,");
            sb.Append("cd_rg_presion, de_rg_presion, cd_metodo_fact, de_metodo_fact, lg_telemedida, cd_tp_gasinera, de_tp_gasinera, fh_desde_reac, fh_hasta_reac, de_marca_back, lg_contrato_simul, cd_id_producto, cd_tp_producto,");
            sb.Append("de_tp_producto, lg_arrastre_penali, nm_med_capacitiva, nm_prec_capacitiva, nm_exceso_pot_p1, nm_exceso_pot_p2, nm_exceso_pot_p3, nm_exceso_pot_p4, nm_exceso_pot_p5, nm_exceso_pot_p6, nm_consumo_medio_5a,");
            sb.Append("nm_consumo_medio");
            sb.Append(" from ed_owner.t_ed_h_facts_atr_ml ml where");            
            sb.Append(" ml.fh_act_dmco >= '").Append(fd.ToString("yyyy-MM-dd")).Append("' AND ");
            sb.Append(" cd_cups in ('").Append(lista_cups13[0]).Append("'");

            for (int i = 1; i < lista_cups13.Count; i++)
                sb.Append(",'").Append(lista_cups13[i]).Append("'");

            sb.Append(")");

            ficheroLog.Add(sb.ToString());
            return sb.ToString();
        }

        private string Consulta_Peajes_cups20(DateTime fd, Dictionary<string, string> dic_cups20)
        {
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;

            sb.Append("SELECT cl_crto_ext_ps, id_crto_ext_ps, id_crto_ps, cd_sec_fact, cd_linea_neg, de_linea_neg, cd_empr_distr, cd_est_fact, cd_mes, id_fact, fh_fact, fh_ini_fact, fh_fin_fact,");
            sb.Append("nm_importe, nm_meses, cd_tp_fact, de_tp_fact, id_fact_sust, cd_cups, cd_cups_ext, cd_empr_distr_cne, de_empr_distr_cne, nm_med_potencia_1, nm_prec_potencia_1, nm_med_potencia_2, nm_prec_potencia_2,");
            sb.Append("nm_med_potencia_3, nm_prec_potencia_3, nm_med_potencia_4, nm_prec_potencia_4, nm_med_potencia_5, nm_prec_potencia_5, nm_med_potencia_6, nm_prec_potencia_6, nm_med_potencia_7, nm_prec_potencia_7,");
            sb.Append("nm_med_potencia_8, nm_prec_potencia_8, nm_med_potencia_9, nm_prec_potencia_9, nm_med_potencia_10, nm_prec_potencia_10, nm_med_activa_1, nm_prec_activa_1, nm_med_activa_2, nm_prec_activa_2,");
            sb.Append("nm_med_activa_3, nm_prec_activa_3, nm_med_activa_4, nm_prec_activa_4, nm_med_activa_5, nm_prec_activa_5, nm_med_activa_6, nm_prec_activa_6, nm_med_activa_7, nm_prec_activa_7, nm_med_activa_8,");
            sb.Append("nm_prec_activa_8, nm_med_activa_9, nm_prec_activa_9, nm_med_activa_10, nm_prec_activa_10, nm_med_reactiva_1, nm_prec_reactiva_1, nm_med_reactiva_2, nm_prec_reactiva_2, nm_med_reactiva_3,");
            sb.Append("nm_prec_reactiva_3, nm_med_reactiva_4, nm_prec_reactiva_4, nm_med_reactiva_5, nm_prec_reactiva_5, nm_med_reactiva_6, nm_prec_reactiva_6, nm_med_reactiva_7, nm_prec_reactiva_7, nm_med_reactiva_8,");
            sb.Append("nm_prec_reactiva_8, nm_med_reactiva_9, nm_prec_reactiva_9, nm_med_reactiva_10, nm_prec_reactiva_10, cd_concepto_1, cd_concepto_1_sce, de_concepto_1, im_concepto_1, cd_concepto_2, cd_concepto_2_sce,");
            sb.Append("de_concepto_2, im_concepto_2, cd_concepto_3, cd_concepto_3_sce, de_concepto_3, im_concepto_3, cd_concepto_4, cd_concepto_4_sce, de_concepto_4, im_concepto_4, cd_concepto_5, cd_concepto_5_sce, de_concepto_5,");
            sb.Append("im_concepto_5, cd_concepto_6, cd_concepto_6_sce, de_concepto_6, im_concepto_6, cd_concepto_7, cd_concepto_7_sce, de_concepto_7, im_concepto_7, cd_concepto_8, cd_concepto_8_sce, de_concepto_8, im_concepto_8,");
            sb.Append("cd_concepto_9, cd_concepto_9_sce, de_concepto_9, im_concepto_9, cd_concepto_10, cd_concepto_10_sce, de_concepto_10, im_concepto_10, cd_concepto_11, cd_concepto_11_sce, de_concepto_11, im_concepto_11,");
            sb.Append("cd_concepto_12, cd_concepto_12_sce, de_concepto_12, im_concepto_12, cd_concepto_13, cd_concepto_13_sce, de_concepto_13, im_concepto_13, cd_concepto_14, cd_concepto_14_sce, de_concepto_14, im_concepto_14,");
            sb.Append("cd_concepto_15, cd_concepto_15_sce, de_concepto_15, im_concepto_15, cd_concepto_16, cd_concepto_16_sce, de_concepto_16, im_concepto_16, cd_concepto_17, cd_concepto_17_sce, de_concepto_17, im_concepto_17,");
            sb.Append("cd_concepto_18, cd_concepto_18_sce, de_concepto_18, im_concepto_18, cd_concepto_19, cd_concepto_19_sce, de_concepto_19, im_concepto_19, cd_concepto_20, cd_concepto_20_sce, de_concepto_20, im_concepto_20,");
            sb.Append("de_impuesto_1, nm_porcentaje_1, nm_base_1, nm_importe_1, de_impuesto_2, nm_porcentaje_2, nm_base_2, nm_importe_2, de_impuesto_3, nm_porcentaje_3, nm_base_3, nm_importe_3, de_impuesto_4, nm_porcentaje_4,");
            sb.Append("nm_base_4, nm_importe_4, de_impuesto_5, nm_porcentaje_5, nm_base_5, nm_importe_5, cod_carga_ods, fh_act_ods, fh_act_dmco, cd_tarifa, cd_tarifa_ff, nm_pot_ctatada, cd_comercializadora, cd_tipo_consumo,");
            sb.Append("cd_municipio, cd_cups20_metra, cd_tp_teleg, consumo_tot_act, consumo_punta, consumo_llano, consumo_valle, consumo_activa4, consumo_activa5, consumo_activa6, consumo_tot_react, consumo_reactiva1, consumo_reactiva2,");
            sb.Append("consumo_reactiva3, consumo_reactiva4, consumo_reactiva5, consumo_reactiva6, potencia_maxpunta, potencia_maxllano, potencia_maxvalle, potencia_max4, potencia_max5, potencia_max6, consumo_supervalle, fh_recepcion,");
            sb.Append("cod_carga, id_pseudofactura, cd_comercializadora_fact, cd_empr_comer_cne, id_fact_recti, id_nif_dist, de_comcnmc, im_salfact, tp_moneda, fh_feboe, lg_med_perdida, nm_vastrafo, nm_porc_perdida, cd_fact_curva,");
            sb.Append("fh_desde_cch, fh_hasta_cch, nm_num_mes, fh_lect_ant_punta, cd_tp_fuente_ant_punta, nm_lect_ant_punta, fh_lect_act_punta, cd_tp_fuente_act_punta, nm_lect_act_punta, im_ajuste_integr_punta, cd_tp_anoma_punta,");
            sb.Append("fh_lect_ant_llano, cd_tp_fuente_ant_llano, nm_lect_ant_llano, fh_lect_act_llano, cd_tp_fuente_act_llano, nm_lect_act_llano, im_ajuste_integr_llano, cd_tp_anoma_llano, fh_lect_ant_valle, cd_tp_fuente_ant_valle,");
            sb.Append("nm_lect_ant_valle, fh_lect_act_valle, cd_tp_fuente_act_valle, nm_lect_act_valle, im_ajuste_integr_valle, cd_tp_anoma_valle, fh_lect_ant_activa4, cd_tp_fuente_ant_activa4, nm_lect_ant_activa4, fh_lect_act_activa4,");
            sb.Append("cd_tp_fuente_act_activa4, nm_lect_act_activa4, im_ajuste_integr_activa4, cd_tp_anoma_activa4, fh_lect_ant_activa5, cd_tp_fuente_ant_activa5, nm_lect_ant_activa5, fh_lect_act_activa5, cd_tp_fuente_act_activa5,");
            sb.Append("nm_lect_act_activa5, im_ajuste_integr_activa5, cd_tp_anoma_activa5, fh_lect_ant_activa6, cd_tp_fuente_ant_activa6, nm_lect_ant_activa6, fh_lect_act_activa6, cd_tp_fuente_act_activa6, nm_lect_act_activa6,");
            sb.Append("im_ajuste_integr_activa6, cd_tp_anoma_activa6, fh_lect_ant_reactiva1, cd_tp_fuente_ant_reactiva1, nm_lect_ant_reactiva1, fh_lect_act_reactiva1, cd_tp_fuente_act_reactiva1, nm_lect_act_reactiva1,");
            sb.Append("im_ajuste_integr_reactiva1, cd_tp_anoma_reactiva1, fh_lect_ant_reactiva2, cd_tp_fuente_ant_reactiva2, nm_lect_ant_reactiva2, fh_lect_act_reactiva2, cd_tp_fuente_act_reactiva2, nm_lect_act_reactiva2,");
            sb.Append("im_ajuste_integr_reactiva2, cd_tp_anoma_reactiva2, fh_lect_ant_reactiva3, cd_tp_fuente_ant_reactiva3, nm_lect_ant_reactiva3, fh_lect_act_reactiva3, cd_tp_fuente_act_reactiva3, nm_lect_act_reactiva3,");
            sb.Append("im_ajuste_integr_reactiva3, cd_tp_anoma_reactiva3, fh_lect_ant_reactiva4, cd_tp_fuente_ant_reactiva4, nm_lect_ant_reactiva4, fh_lect_act_reactiva4, cd_tp_fuente_act_reactiva4, nm_lect_act_reactiva4,");
            sb.Append("im_ajuste_integr_reactiva4, cd_tp_anoma_reactiva4, fh_lect_ant_reactiva5, cd_tp_fuente_ant_reactiva5, nm_lect_ant_reactiva5, fh_lect_act_reactiva5, cd_tp_fuente_act_reactiva5, nm_lect_act_reactiva5,");
            sb.Append("im_ajuste_integr_reactiva5, cd_tp_anoma_reactiva5, fh_lect_ant_reactiva6, cd_tp_fuente_ant_reactiva6, nm_lect_ant_reactiva6, fh_lect_act_reactiva6, cd_tp_fuente_act_reactiva6, nm_lect_act_reactiva6,");
            sb.Append("im_ajuste_integr_reactiva6, cd_tp_anoma_reactiva6, fh_lect_ant_maxpunta, cd_tp_fuente_ant_maxpunta, nm_lect_ant_maxpunta, fh_lect_act_maxpunta, cd_tp_fuente_act_maxpunta, nm_lect_act_maxpunta,");
            sb.Append("im_ajuste_integr_maxpunta, cd_tp_anoma_maxpunta, fh_lect_ant_maxllano, cd_tp_fuente_ant_maxllano, nm_lect_ant_maxllano, fh_lect_act_maxllano, cd_tp_fuente_act_maxllano, nm_lect_act_maxllano,");
            sb.Append("im_ajuste_integr_maxllano, cd_tp_anoma_maxllano, fh_lect_ant_maxvalle, cd_tp_fuente_ant_maxvalle, nm_lect_ant_maxvalle, fh_lect_act_maxvalle, cd_tp_fuente_act_maxvalle, nm_lect_act_maxvalle,");
            sb.Append("im_ajuste_integr_maxvalle, cd_tp_anoma_maxvalle, fh_lect_ant_max4, cd_tp_fuente_ant_max4, nm_lect_ant_max4, fh_lect_act_max4, cd_tp_fuente_act_max4, nm_lect_act_max4, im_ajuste_integr_max4,");
            sb.Append("cd_tp_anoma_max4, fh_lect_ant_max5, cd_tp_fuente_ant_max5, nm_lect_ant_max5, fh_lect_act_max5, cd_tp_fuente_act_max5, nm_lect_act_max5, im_ajuste_integr_max5, cd_tp_anoma_max5, fh_lect_ant_max6,");
            sb.Append("cd_tp_fuente_ant_max6, nm_lect_ant_max6, fh_lect_act_max6, cd_tp_fuente_act_max6, nm_lect_act_max6, im_ajuste_integr_max6, cd_tp_anoma_max6, fh_lect_ant_supervalle, cd_tp_fuente_ant_supervalle,");
            sb.Append("nm_lect_ant_supervalle, fh_lect_act_supervalle, cd_tp_fuente_act_supervalle, nm_lect_act_supervalle, im_ajuste_integr_supervalle, cd_tp_anoma_supervalle, cd_concepto_reper_1, im_concepto_reper_1,");
            sb.Append("cd_concepto_reper_2, im_concepto_reper_2, cd_concepto_reper_3, im_concepto_reper_3, cd_concepto_reper_4, im_concepto_reper_4, cd_concepto_reper_5, im_concepto_reper_5, cd_concepto_reper_6, im_concepto_reper_6,");
            sb.Append("cd_concepto_reper_7, im_concepto_reper_7, cd_concepto_reper_8, im_concepto_reper_8, cd_concepto_reper_9, im_concepto_reper_9, cd_concepto_reper_10, im_concepto_reper_10, cd_concepto_reper_11, im_concepto_reper_11,");
            sb.Append("cd_concepto_reper_12, im_concepto_reper_12, cd_cmunicip_atr, de_est_fact, fh_desde_ener, fh_hasta_ener, fh_desde_pot, fh_hasta_pot, cd_modo_crtl_pot, nm_dias_fact, im_tot_fact, im_saldo_tot_fact, nm_tot_recibos,");
            sb.Append("fh_valor, fh_lim_pago, id_remesa, cd_munic_red, lg_penaliza_icp, identificador, cd_mensaje, cd_doc_ref, cd_tp_doc, de_tp_doc, de_tp_aceso_serv, lg_relevancia_recl_pdp, de_empr_comer_cne, cd_huella_fact_atr, fichero,");
            sb.Append("cd_huella_blq_fich, fh_publicacion, fh_public_fichero, cd_solicitud, cd_sec_solicitud, lg_atr, cd_tp_reg, de_tp_reg, cd_expdte, lg_bloq_af, lg_recl_abta, cd_origen_fact, de_origen_fact, cd_aniofactura, cd_pais,");
            sb.Append("lg_autoconsumo, cd_tp_pm, lg_duracioninfanio, fh_creacion, cd_usuario_creador, de_usuario_creador, fh_mod, cd_usuario_mod, de_usuario_mod, cd_ubicacion, de_ubicacion, id_exp_af, lg_registro_borrado, de_tipo_borrado,");
            sb.Append("cd_rg_presion, de_rg_presion, cd_metodo_fact, de_metodo_fact, lg_telemedida, cd_tp_gasinera, de_tp_gasinera, fh_desde_reac, fh_hasta_reac, de_marca_back, lg_contrato_simul, cd_id_producto, cd_tp_producto,");
            sb.Append("de_tp_producto, lg_arrastre_penali, nm_med_capacitiva, nm_prec_capacitiva, nm_exceso_pot_p1, nm_exceso_pot_p2, nm_exceso_pot_p3, nm_exceso_pot_p4, nm_exceso_pot_p5, nm_exceso_pot_p6, nm_consumo_medio_5a,");
            sb.Append("nm_consumo_medio");
            sb.Append(" from ed_owner.t_ed_h_facts_atr_ml ml where");
            sb.Append(" ml.fh_act_dmco >= '").Append(fd.AddDays(-1).ToString("yyyy-MM-dd")).Append("' AND ");

            foreach(KeyValuePair<string, string> p in dic_cups20)
            {
                if (firstOnly)
                {
                    sb.Append(" cd_cups20_metra in ('").Append(p.Key).Append("'");
                    firstOnly = false;
                }else
                    sb.Append(",'").Append(p.Key).Append("'");
            }
                      

            sb.Append(")");

            ficheroLog.Add(sb.ToString());
            return sb.ToString();
        }

        private string Consulta_Count(DateTime fd, string cups13)
        {
            StringBuilder sb = new StringBuilder();


            sb.Append("SELECT count(*) total_registos");
            sb.Append(" from ed_owner.t_ed_h_facts_atr_ml ml where");
            sb.Append(" cd_cups = '" + cups13 + "' AND");
            sb.Append(" ml.fh_act_dmco >= '").Append(fd.ToString("yyyy-MM-dd")).Append("'");

            return sb.ToString();
        }

        private List<EndesaEntity.medida.Peajes_Vista> Consulta_Peajes_MySQL(List<EndesaEntity.medida.PuntoSuministro> lista_cups, bool fecha_consumo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            bool firstOny = true;

            List<EndesaEntity.medida.Peajes_Vista> lista = new List<EndesaEntity.medida.Peajes_Vista>();
            DateTime fecha_min = new DateTime();
            DateTime fecha_max = new DateTime();


            try
            {
                fecha_min = new DateTime(4999, 12, 31);
                fecha_max = DateTime.MinValue;

                // Buscamos las fechas máximas y mínimas para la consulta
                foreach (EndesaEntity.medida.PuntoSuministro p in lista_cups)
                {
                    if (p.fd < fecha_min)
                        fecha_min = p.fd;

                    if (p.fh > fecha_max)
                        fecha_max = p.fh;

                }

                strSql = "SELECT CUPS13, CUPS20, Cod_Factura, Estado_Factura, Tipo_Factura,"
                    + " Fecha_Factura, Fecha_Desde, Fecha_Hasta, Ind_Cuadre_Periodo_c, FFM_AE_total_c, FFM_R1_total_c,"
                    + " FFM_PMax_total_c, FFF_AE_total_c, FFF_R1_total_c, FFF_PFac_total_c, FFF_Imp_Exc_Pot_total_c,"
                    + " FFF_Imp_Exc_Pot, FFF_Imp_Exc_R1, FFM_AE_1, FFM_AE_2, FFM_AE_3, FFM_AE_4, FFM_AE_5, FFM_AE_6,"
                    + " FFM_R1_1, FFM_R1_2, FFM_R1_3, FFM_R1_4, FFM_R1_5, FFM_R1_6,"
                    + " FFM_PMax_1, FFM_PMax_2, FFM_PMax_3, FFM_PMax_4, FFM_PMax_5, FFM_PMax_6,"
                    + " FFF_AE_1, FFF_AE_2, FFF_AE_3, FFF_AE_4, FFF_AE_5, FFF_AE_6,"
                    + " FFF_R1_1, FFF_R1_2, FFF_R1_3, FFF_R1_4, FFF_R1_5, FFF_R1_6, FFF_R4_6,"
                    + " FFF_PFac1, FFF_PFac2, FFF_PFac3, FFF_PFac4, FFF_PFac5, FFF_PFac6,"
                    + " FFF_Imp_Exc_Pot_1, FFF_Imp_Exc_Pot_2, FFF_Imp_Exc_Pot_3,"
                    + " FFF_Imp_Exc_Pot_4, FFF_Imp_Exc_Pot_5, FFF_Imp_Exc_Pot_6,"
                    + " CUPS20_Metra, Procedencia, Ind_Perdidas, Porcentaje_Perdidas,"
                    + " Pot_Trafo_Perdidas_VA, Tipo_PM, Ind_Autoconsumo, Ind_Telemedida,"
                    + " Cod_Factura_Sustituida, Cod_Factura_Rectificada, Codigo_Estado_Factura,"
                    + " Tarifa, Tipo_Consumo, Tipo_Telegestion, Contrato_Ext_PS, Contrato_PS, Sec_Factura,"
                    + " Distribuidora, Consumo_Tot_Act, Consumo_Tot_React, Fecha_desde_AE, Fecha_hasta_AE,"
                    + " Fecha_desde_R1, Fecha_hasta_R1, Fecha_desde_PFac, Fecha_hasta_PFac, Fecha_desde_Curva,"
                    + " Fecha_hasta_Curva, CUPS_EXT, Cod_Carga_ODS, Fecha_act_ODS, Fecha_act_DMCO, Fecha_Recepcion,"
                    + " Cod_Carga"
                    + " FROM med.t_ed_h_facts_atr_ml_v WHERE";

                foreach(EndesaEntity.medida.PuntoSuministro p in lista_cups)
                {
                    if (firstOny)
                    {
                        strSql += " CUPS20 in ('" + p.cups20 + "'";
                        firstOny = false;
                    }else
                        strSql += ",'" + p.cups20 + "'";

                }
                strSql += ") and ";

                if (fecha_consumo)
                {
                    strSql += "(Fecha_Desde >= '" + fecha_min.ToString("yyyy-MM-dd") + "' and"
                        + " Fecha_Hasta <= '" + fecha_max.ToString("yyyy-MM-dd") + "')";
                }
                else
                {
                    strSql += "(Fecha_Factura >= '" + fecha_min.ToString("yyyy-MM-dd") + "' and"
                        + " Fecha_Factura <= '" + fecha_max.ToString("yyyy-MM-dd") + "')";
                }


                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Peajes_Vista c = new EndesaEntity.medida.Peajes_Vista();
                    if (r["CUPS13"] != System.DBNull.Value)
                        c.cups13 = r["CUPS13"].ToString();

                    if (r["CUPS20"] != System.DBNull.Value)
                        c.cups20 = r["CUPS20"].ToString();

                    if (r["Cod_Factura"] != System.DBNull.Value)
                        c.cod_factura = r["Cod_Factura"].ToString();

                    if (r["Estado_Factura"] != System.DBNull.Value)
                        c.estado_factura = r["Estado_Factura"].ToString();

                    if (r["Tipo_Factura"] != System.DBNull.Value)
                        c.tipo_factura = r["Tipo_Factura"].ToString();

                    if (r["Fecha_Factura"] != System.DBNull.Value)
                        c.fecha_factura = Convert.ToDateTime(r["Fecha_Factura"]);

                    if (r["Fecha_Desde"] != System.DBNull.Value)
                        c.fecha_desde = Convert.ToDateTime(r["Fecha_Desde"]);

                    if (r["Fecha_Hasta"] != System.DBNull.Value)
                        c.fecha_hasta = Convert.ToDateTime(r["Fecha_Desde"]);

                    if (r["Ind_Cuadre_Periodo_c"] != System.DBNull.Value)
                        c.ind_cuadre_periodo_c = r["Ind_Cuadre_Periodo_c"].ToString();

                    if (r["FFM_AE_total_c"] != System.DBNull.Value)
                        c.ffm_ae_total_c = Convert.ToDouble(r["FFM_AE_total_c"]);

                    if (r["FFM_R1_total_c"] != System.DBNull.Value)
                        c.ffm_r1_total_c = Convert.ToDouble(r["FFM_R1_total_c"]);

                    if (r["FFM_PMax_total_c"] != System.DBNull.Value)
                        c.ffm_pmax_total_c = Convert.ToDouble(r["FFM_PMax_total_c"]);

                    if (r["FFF_AE_total_c"] != System.DBNull.Value)
                        c.fff_pfac_total_c = Convert.ToDouble(r["FFF_AE_total_c"]);

                    if (r["FFF_R1_total_c"] != System.DBNull.Value)
                        c.fff_r1_total_c = Convert.ToDouble(r["FFF_R1_total_c"]);

                    if (r["FFF_PFac_total_c"] != System.DBNull.Value)
                        c.fff_pfac_total_c = Convert.ToDouble(r["FFF_PFac_total_c"]);

                    if (r["FFF_Imp_Exc_Pot_total_c"] != System.DBNull.Value)
                        c.fff_imp_exc_pot_total_c = Convert.ToDouble(r["FFF_Imp_Exc_Pot_total_c"]);

                    if (r["FFF_Imp_Exc_Pot"] != System.DBNull.Value)
                        c.fff_imp_exc_pot = Convert.ToDouble(r["FFF_Imp_Exc_Pot"]);

                    if (r["FFF_Imp_Exc_R1"] != System.DBNull.Value)
                        c.fff_imp_exc_r1 = Convert.ToDouble(r["FFF_Imp_Exc_R1"]);                   

                    for (int i = 1; i <= 6; i++)
                        if (r["FFM_AE_" + i] != System.DBNull.Value)
                            c.ffm_ae[i - 1] = Convert.ToDouble(r["FFM_AE_" + i]);

                    for (int i = 1; i <= 6; i++)
                        if (r["FFM_R1_" + i] != System.DBNull.Value)
                            c.ffm_r1[i - 1] = Convert.ToDouble(r["FFM_R1_" + i]);

                    if (r["FFF_R4_6"] != System.DBNull.Value)
                        c.fff_r4_6 = Convert.ToDouble(r["FFF_R4_6"]);

                    for (int i = 1; i <= 6; i++)
                        if (r["FFM_PMax_" + i] != System.DBNull.Value)
                            c.ffm_pmax[i - 1] = Convert.ToDouble(r["FFM_PMax_" + i]);

                    for (int i = 1; i <= 6; i++)
                        if (r["FFF_PFac" + i] != System.DBNull.Value)
                            c.fff_pfac[i - 1] = Convert.ToDouble(r["FFF_PFac" + i]);

                    for (int i = 1; i <= 6; i++)
                        if (r["FFF_Imp_Exc_Pot_" + i] != System.DBNull.Value)
                            c.fff_imp_exc_pot1[i - 1] = Convert.ToDouble(r["FFF_Imp_Exc_Pot_" + i]);                     
                   

                    if (r["CUPS20_Metra"] != System.DBNull.Value)
                        c.cups20_metra = r["CUPS20_Metra"].ToString();

                    if (r["Procedencia"] != System.DBNull.Value)
                        c.procedencia = r["Procedencia"].ToString();

                    if (r["Ind_Perdidas"] != System.DBNull.Value)
                        c.ind_perdidas = r["Ind_Perdidas"].ToString();

                    if (r["Porcentaje_Perdidas"] != System.DBNull.Value)
                        c.porcentaje_perdidas = Convert.ToDouble(r["Porcentaje_Perdidas"]);

                    if (r["Pot_Trafo_Perdidas_VA"] != System.DBNull.Value)
                        c.pot_trafo_perdidas_va = Convert.ToDouble(r["Pot_Trafo_Perdidas_VA"]);

                    if (r["Tipo_PM"] != System.DBNull.Value)
                        c.tipo_pm = r["Tipo_PM"].ToString();

                    if (r["Ind_Autoconsumo"] != System.DBNull.Value)
                        c.ind_autoconsumo = r["Ind_Autoconsumo"].ToString();

                    if (r["Ind_Telemedida"] != System.DBNull.Value)
                        c.ind_telemedida = r["Ind_Telemedida"].ToString();                    

                    if (r["Cod_Factura_Sustituida"] != System.DBNull.Value)
                        c.cod_factura_sustituida = r["Cod_Factura_Sustituida"].ToString();

                    if (r["Cod_Factura_Rectificada"] != System.DBNull.Value)
                        c.cod_factura_rectificada = r["Cod_Factura_Rectificada"].ToString();

                    if (r["Codigo_Estado_Factura"] != System.DBNull.Value)
                        c.codigo_estado_factura = r["Codigo_Estado_Factura"].ToString();

                    if (r["Tarifa"] != System.DBNull.Value)
                        c.tarifa = r["Tarifa"].ToString();

                    if (r["Tipo_Consumo"] != System.DBNull.Value)
                        c.tipo_consumo = r["Tipo_Consumo"].ToString();

                    if (r["Tipo_Telegestion"] != System.DBNull.Value)
                        c.tipo_telegestion = r["Tipo_Telegestion"].ToString();

                    if (r["Contrato_Ext_PS"] != System.DBNull.Value)
                        c.contrato_ext_ps = r["Contrato_Ext_PS"].ToString();

                    if (r["Contrato_PS"] != System.DBNull.Value)
                        c.contrato_ps = r["Contrato_PS"].ToString();

                    if (r["Sec_Factura"] != System.DBNull.Value)
                        c.sec_factura = Convert.ToInt32(r["Sec_Factura"]);

                    if (r["Distribuidora"] != System.DBNull.Value)
                        c.distribuidora = r["Distribuidora"].ToString();                    

                    if (r["Consumo_Tot_Act"] != System.DBNull.Value)
                        c.consumo_tot_act = Convert.ToDouble(r["Consumo_Tot_Act"]);

                    if (r["Consumo_Tot_React"] != System.DBNull.Value)
                        c.consumo_tot_react = Convert.ToDouble(r["Consumo_Tot_React"]);

                    
                    
                    //+ " CUPS_EXT, Cod_Carga_ODS, Fecha_act_ODS, Fecha_act_DMCO, Fecha_Recepcion,"
                    //+ " Cod_Carga"

                    if (r["Fecha_desde_AE"] != System.DBNull.Value)
                        c.fecha_desde_ae = Convert.ToDateTime(r["Fecha_desde_AE"]);

                    if (r["Fecha_hasta_AE"] != System.DBNull.Value)
                        c.fecha_hasta_ae = Convert.ToDateTime(r["Fecha_hasta_AE"]);

                    if (r["Fecha_desde_R1"] != System.DBNull.Value)
                        c.fecha_desde_r1 = Convert.ToDateTime(r["Fecha_desde_R1"]);

                    if (r["Fecha_hasta_R1"] != System.DBNull.Value)
                        c.fecha_hasta_r1 = Convert.ToDateTime(r["Fecha_hasta_R1"]);

                    if (r["Fecha_desde_PFac"] != System.DBNull.Value)
                        c.fecha_desde_pfac = Convert.ToDateTime(r["Fecha_desde_PFac"]);

                    if (r["Fecha_hasta_PFac"] != System.DBNull.Value)
                        c.fecha_hasta_pfac = Convert.ToDateTime(r["Fecha_hasta_PFac"]);

                    if (r["Fecha_desde_Curva"] != System.DBNull.Value)
                        c.fecha_desde_curva = Convert.ToDateTime(r["Fecha_desde_Curva"]);

                    if (r["Fecha_hasta_Curva"] != System.DBNull.Value)
                        c.fecha_hasta_curva = Convert.ToDateTime(r["Fecha_hasta_Curva"]);

                    if (r["CUPS_EXT"] != System.DBNull.Value)
                        c.cups_ext = r["CUPS_EXT"].ToString();

                    if (r["Cod_Carga_ODS"] != System.DBNull.Value)
                        c.cod_carga_ods = Convert.ToInt32(r["Cod_Carga_ODS"]);

                    if (r["Fecha_act_ODS"] != System.DBNull.Value)
                        c.fecha_act_ods = Convert.ToDateTime(r["Fecha_act_ODS"]);

                    if (r["Fecha_Recepcion"] != System.DBNull.Value)
                        c.fecha_recepcion = Convert.ToDateTime(r["Fecha_Recepcion"]);

                    if (r["Cod_Carga"] != System.DBNull.Value)
                        c.cod_carga = Convert.ToInt32(r["Cod_Carga"]);


                    if (fecha_consumo)
                    {
                        if (lista_cups.Exists(z => z.cups20 == c.cups20 && (z.fd <= c.fecha_desde && z.fh >= c.fecha_hasta)))
                            lista.Add(c);
                    }
                    else
                    {
                        if (lista_cups.Exists(z => z.cups20 == c.cups20 && (z.fd <= c.fecha_factura && z.fh >= c.fecha_factura)))
                            lista.Add(c);
                    }

                    


                }
                db.CloseConnection();

                return lista;

            }catch(Exception ex)
            {
                return null;
            }
        }



        private void PintaRecuadro(ExcelPackage excelPackage, int f, int c)
        {
            var workSheet = excelPackage.Workbook.Worksheets.First();
            workSheet.Cells[f, c].Style.Border.Top.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[f, c].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[f, c].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
        }
    }
}
