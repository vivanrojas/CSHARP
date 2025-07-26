using EndesaBusiness.servidores;
using EndesaEntity;
using Microsoft.Graph.TermStore;
using Microsoft.Office.Interop.Access;
using Microsoft.SharePoint.Client;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using Renci.SshNet.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Telegram.Bot.Types;

namespace EndesaBusiness.facturacion
{
    public class Tunel
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "Tunel");
        Dictionary<string, List<EndesaEntity.Tunel>> dic;
        List<EndesaEntity.Tunel> lista;
        EndesaBusiness.utilidades.Param p;

        public Tunel()
        {
            p = new EndesaBusiness.utilidades.Param("tunel_param", MySQLDB.Esquemas.FAC);
        }
                

        public void GeneraCuadroMando()
        {
            if(p.GetValue("ejecuta_proceso") == "S")
            {
                Console.WriteLine("Lanzamos la carga de contratos");
                Console.WriteLine("==============================");
                lista = CargaTablaContratos();
                ProcesoCuadroMando();
            }            
        }


        public void CargaExcel(string fichero, bool mensual)
        {
            dic = ProcesaExcel(fichero, mensual);
            RellenaInventario(dic);
            Proceso(mensual);
            GeneraExcelResultados(fichero);
        }

        private List<EndesaEntity.Tunel> CargaTablaContratos()
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            
            List<EndesaEntity.Tunel> l = new List<EndesaEntity.Tunel>();

            try
            {


                strSql = "SELECT NAME, version_start, version_end, volume,"
                     + " lowerband, higherband, start_date, end_date, cups20"
                     + " FROM tunel_contratos c";
                Console.WriteLine(strSql);
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.Tunel t = new EndesaEntity.Tunel();                    
                    t.name = r["NAME"].ToString();
                    t.versionStart = Convert.ToDateTime(r["version_start"]);
                    t.versionEnd = Convert.ToDateTime(r["version_end"]); 
                    t.volume = Convert.ToDouble(r["volume"]);
                    t.lowerBand = Convert.ToDouble(r["lowerband"]); 
                    t.higherBand = Convert.ToDouble(r["higherband"]);
                    t.startDate = Convert.ToDateTime(r["start_date"]);
                    t.endDate = Convert.ToDateTime(r["end_date"]);
                    t.cups20 = Convert.ToString(r["cups20"]); 
                    //t.cups13 = Convert.ToString(r[""]);
                    //t.tarifa = Convert.ToString(r[""]);
                    l.Add(t);
                }
                db.CloseConnection();
                return l;
            }
            catch (Exception e)
            {
                
                MessageBox.Show(e.Message,
                    "Estructura de columnas Excel incorrecta",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return null;
            }

        }

        private void ProcesoCuadroMando()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string cups20 = "";
            

            

            List<EndesaEntity.Tunel> lista_totales = new List<EndesaEntity.Tunel>();

            try
            {
                //strSql = "DELETE FROM tunel_facturas_tmp";
                //Console.WriteLine(strSql);
                //db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //strSql = "DELETE FROM tunel_facturas";
                //Console.WriteLine(strSql);
                //db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //strSql = "REPLACE INTO tunel_facturas_tmp"
                //    + " SELECT"
                //    + " f.CEMPTITU, f.CNIFDNIC, f.DAPERSOC,"
                //    + " f.CREFEREN, f.SECFACTU, f.TFACTURA,"
                //    + " f.TESTFACT, f.CFACTURA, f.CFACTREC,"
                //    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                //    + " f.CUPSREE, f.IFACTURA, f.VCUOVAFA"
                //    + " FROM fact.tunel_contratos c"
                //    + " INNER JOIN fact.fo f ON"
                //    + " (f.FFACTDES >= c.start_date and f.FFACTHAS <= c.end_date)"
                //    + " and substr(f.CUPSREE, 1, 20) = c.cups20";
                    
                //Console.WriteLine(strSql);
                //db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //strSql = "delete from tunel_facturas_tmp where TESTFACT in ('A','S')";
                //Console.WriteLine(strSql);
                //db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //strSql = "DELETE FROM tunel_facturas_tmp WHERE VCUOVAFA = 0";
                //Console.WriteLine(strSql);
                //db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //strSql = "REPLACE INTO tunel_facturas"
                //    + " SELECT * FROM tunel_facturas_tmp f"
                //    + " ORDER BY f.CEMPTITU, f.CNIFDNIC, f.CREFEREN,"
                //    + " f.CUPSREE, f.FFACTDES, f.SECFACTU";
                //Console.WriteLine(strSql);
                //db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //strSql = "UPDATE tunel_contratos c"
                //     + " INNER JOIN"
                //     + " (SELECT substr(f.CUPSREE, 1, 20) AS CUPS20,"
                //     + " SUM(f.VCUOVAFA) AS consumo_total,"
                //     + " MAX(f.FFACTHAS) AS max_fecha" 
                //     + " FROM tunel_facturas f GROUP BY f.CUPSREE) AS f on"
                //     + " f.CUPS20 = c.cups20"
                //     + " SET c.fecha_ultima_factura = f.max_fecha,"
	               //  + " c.consumo_total = f.consumo_total";
                //Console.WriteLine(strSql);
                //db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();


                strSql = "UPDATE tunel_contratos c"
                    + " INNER JOIN"
                    + " (select"
                    + " substr(ff.CUPSREE, 1, 20) AS CUPS20,"
                    + "     MIN(ff.FFACTDES) as fecha_primera_factura,"  
                    + "     MAX(ff.FFACTHAS) AS fecha_ultima_factura,"
                    + " SUM(ff.VCUOVAFA) AS consumo_total"
                    + "     FROM tunel_contratos cc"
                    + "     INNER JOIN fo ff ON"
                    + "     substr(ff.CUPSREE, 1, 20) = cc.cups20"
                    + "     AND(ff.FFACTDES >= cc.start_date and ff.FFACTHAS <= cc.end_date)"
                    + "     GROUP BY ff.CUPSREE) AS f on"
                    + "     f.CUPS20 = c.cups20"
                    + "     SET c.fecha_ultima_factura = f.fecha_ultima_factura,"
                    + "     c.fecha_primera_factura = f.fecha_primera_factura,"
                    + "     c.consumo_total = f.consumo_total";

                Console.WriteLine(strSql);
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                GeneraResumen();


                // Actualizamos la fecha de ejecucion del proceso
                p.SetCodeValueToDateTime("ultima_ejecucion_proceso");

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "Tunel.Proceso",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }

        private void GeneraResumen()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;


        }


        private Dictionary<string, List<EndesaEntity.Tunel>> ProcesaExcel(string fichero, bool mensual)
        {
            Dictionary<string, List<EndesaEntity.Tunel>> d
                = new Dictionary<string, List<EndesaEntity.Tunel>>();
                        
            int c = 1;
            int f = 1;
            int total_hojas_excel = 0;
            bool firstOnly = true;
            string cabecera = "";
            FileStream fs;
            ExcelPackage excelPackage;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            List<EndesaEntity.Tunel> lista = new List<EndesaEntity.Tunel>();
            int dias_del_mes = 0;
            DateTime fecha_inicio_mensual = new DateTime();
            DateTime fecha_fin_mensual = new DateTime();

            try
            {                
                fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                excelPackage = new ExcelPackage(fs);
                total_hojas_excel = excelPackage.Workbook.Worksheets.Count();
                var workSheet = excelPackage.Workbook.Worksheets.First();

                for (int hoja = 0; hoja < total_hojas_excel; hoja++)
                {
                    workSheet = excelPackage.Workbook.Worksheets[hoja];
                    firstOnly = true;

                    if (workSheet.Cells[1, 1].Value == null)
                        break;
                    cabecera = "";
                    for (int ii = 1; ii <= 12; ii++)
                        cabecera += workSheet.Cells[1, ii].Value.ToString();

                    if (cabecera != p.GetValue("cabecera_excel", DateTime.Now, DateTime.Now))
                    {
                        MessageBox.Show("Estructura de columnas en hoja Excel incorrecta."
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "Las columnas del archivo excel no son las esperadas o están en lugares distintos a los esperados.",
                       "Estructura de columnas Excel incorrecta",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Error);
                        break;
                    }

                    lista = new List<EndesaEntity.Tunel>();

                    f = 1; // Porque la primera fila es la cabecera
                    for (int i = 1; i < 5000; i++)
                    {
                        c = 1;
                        f++;

                        if (workSheet.Cells[f, 1].Value == null)
                            break;

                        if (workSheet.Cells[f, 1].Value.ToString() == "")
                            break;

                        EndesaEntity.Tunel t = new EndesaEntity.Tunel();
                        t.id = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                        t.name = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                        t.versionStart = Convert.ToDateTime(workSheet.Cells[f, c].Value); c++;
                        t.versionEnd = Convert.ToDateTime(workSheet.Cells[f, c].Value); c++;
                        t.volume = Convert.ToDouble(workSheet.Cells[f, c].Value); c++;
                        t.lowerBand = Convert.ToDouble(workSheet.Cells[f, c].Value); c++;
                        t.higherBand = Convert.ToDouble(workSheet.Cells[f, c].Value); c++;
                        t.startDate = Convert.ToDateTime(workSheet.Cells[f, c].Value); c++;
                        t.endDate = Convert.ToDateTime(workSheet.Cells[f, c].Value); c++;
                        t.cups20 = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                        t.cups13 = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                        t.tarifa = Convert.ToString(workSheet.Cells[f, c].Value);
                        lista.Add(t);

                        if (mensual)
                        {
                            fecha_fin_mensual = t.endDate;
                            fecha_inicio_mensual = t.startDate;
                            for (DateTime mes = fecha_inicio_mensual; mes <= fecha_fin_mensual; mes = mes.AddMonths(1))
                            {
                                dias_del_mes = DateTime.DaysInMonth(mes.Year, mes.Month);
                                
                                EndesaEntity.Tunel tt = new EndesaEntity.Tunel();
                                tt.id = t.id;
                                tt.name = t.name;
                                tt.versionStart = t.versionStart;
                                tt.versionEnd = t.versionEnd;
                                tt.volume = t.volume;
                                tt.lowerBand = t.lowerBand;
                                tt.higherBand = t.higherBand;
                                tt.cups20 = t.cups20;
                                tt.cups13 = t.cups13;
                                tt.tarifa = t.tarifa;
                                tt.startDate = mes;
                                tt.endDate= new DateTime(mes.Year, mes.Month, dias_del_mes);

                                lista.Add(tt);
                            }
                        }
                        
                        
                    }

                    d.Add(workSheet.Name, lista);
                }
                fs = null;
                excelPackage = null;
                return d;
            }
            catch (Exception e)
            {
                fs = null;
                excelPackage = null;

                MessageBox.Show(e.Message,
                    "Estructura de columnas Excel incorrecta",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return null;
            }

        }

        private void Proceso(bool mensual)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string cups20 = "";
            DateTime fecha_desde = new DateTime();
            DateTime fecha_hasta = new DateTime();

            int total_registros_progress_bar = 0;
            int progreso = 0;
            double percent = 0;

            forms.FrmProgressBar pb = new forms.FrmProgressBar();

            List<EndesaEntity.Tunel> lista_totales = new List<EndesaEntity.Tunel>();

            try
            {
                strSql = "DELETE FROM tunel_facturas_tmp";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "DELETE FROM tunel_facturas";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "REPLACE INTO tunel_facturas_tmp"
                    + " SELECT"
                    + " f.CEMPTITU, f.CNIFDNIC, f.DAPERSOC,"
                    + " f.CREFEREN, f.SECFACTU, f.TFACTURA,"
                    + " f.TESTFACT, f.CFACTURA, f.CFACTREC,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " f.CUPSREE, f.IFACTURA, f.VCUOVAFA"
                    + " FROM fact.tunel_inventario i"
                    + " INNER JOIN fact.fo f ON"
                    + " substr(f.CUPSREE, 1, 20) = i.cups20 AND"
                    + " (f.FFACTDES >= i.fecha_desde and f.FFACTHAS <= i.fecha_hasta)";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "delete from tunel_facturas_tmp where TESTFACT in ('A','S')"; 
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "DELETE FROM tunel_facturas_tmp WHERE VCUOVAFA = 0";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "REPLACE INTO tunel_facturas"
                    + " SELECT * FROM tunel_facturas_tmp f"
                    + " ORDER BY f.CEMPTITU, f.CNIFDNIC, f.CREFEREN," 
                    + " f.CUPSREE, f.FFACTDES, f.SECFACTU";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                foreach (KeyValuePair<string, List<EndesaEntity.Tunel>> p in dic)
                {
                    for (int i = 0; i < p.Value.Count(); i++)
                        total_registros_progress_bar++;
                }

                pb.Text = "Buscando consumos";
                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = total_registros_progress_bar;

                strSql = "SELECT f.CEMPTITU, f.CNIFDNIC, f.DAPERSOC,"
                        + " f.CREFEREN, MIN(f.FFACTDES) AS FFACTDES,"
                        + " MAX(f.FFACTHAS) AS FFACTHAS,"
                        + " f.CUPSREE, sum(f.VCUOVAFA) TOTAL_CONSUMO"
                        + " FROM tunel_facturas f"
                        + " GROUP BY f.CUPSREE";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.Tunel c = new EndesaEntity.Tunel();
                    c.cups20 = r["CUPSREE"].ToString().Substring(0, 20);
                    c.startDate = Convert.ToDateTime(r["FFACTDES"]);
                    c.endDate = Convert.ToDateTime(r["FFACTHAS"]);
                    c.total_energia = Convert.ToInt32(r["TOTAL_CONSUMO"]);

                    lista_totales.Add(c);

                    // cups20 = r["CUPSREE"].ToString().Substring(0, 20);
                    // fecha_desde = Convert.ToDateTime(r["FFACTDES"]);
                    // fecha_hasta = Convert.ToDateTime(r["FFACTHAS"]);



                    progreso++;

                    percent = (progreso / Convert.ToDouble(total_registros_progress_bar)) * 100;
                    pb.progressBar.Increment(1);
                    pb.txtDescripcion.Text = "Buscando consumos para el CUPS " + cups20;
                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    pb.Refresh();


                    //foreach(KeyValuePair<string, List<EndesaEntity.Tunel>> p in dic)
                    //{
                    //    for(int i = 0; i < p.Value.Count; i++)
                    //    {
                    //        if (p.Value[i].cups20 == cups20 
                    //            && (p.Value[i].startDate >= fecha_desde && p.Value[i].endDate <= fecha_hasta))
                    //            p.Value[i].total_energia += Convert.ToInt32(r["TOTAL_CONSUMO"]);
                    //    }
                    //}                   
                }
                db.CloseConnection();


                if (mensual)
                {
                    strSql = "SELECT f.CEMPTITU, f.CNIFDNIC, f.DAPERSOC,"
                       + " f.CREFEREN, MIN(f.FFACTDES) AS FFACTDES,"
                       + " MAX(f.FFACTHAS) AS FFACTHAS,"
                       + " f.CUPSREE, sum(f.VCUOVAFA) TOTAL_CONSUMO"
                       + " FROM tunel_facturas f"
                       + " GROUP BY f.CUPSREE, YEAR(f.FFACTDES), MONTH(f.FFACTDES)";


                    db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();

                    while (r.Read())
                    {
                        EndesaEntity.Tunel c = new EndesaEntity.Tunel();
                        c.cups20 = r["CUPSREE"].ToString().Substring(0, 20);
                        c.startDate = Convert.ToDateTime(r["FFACTDES"]);
                        c.endDate = Convert.ToDateTime(r["FFACTHAS"]);
                        c.total_energia = Convert.ToInt32(r["TOTAL_CONSUMO"]);

                        lista_totales.Add(c);
                    }
                    db.CloseConnection();
                }
                   
               

                for(int k = 0; k < lista_totales.Count; k++)
                {
                    foreach (KeyValuePair<string, List<EndesaEntity.Tunel>> p in dic)
                    {
                        for (int i = 0; i < p.Value.Count; i++)
                        {
                            if (p.Value[i].cups20 == lista_totales[k].cups20
                                && (p.Value[i].startDate == lista_totales[k].startDate && p.Value[i].endDate == lista_totales[k].endDate))
                                p.Value[i].total_energia = lista_totales[k].total_energia;
                        }
                    }
                }

                pb.Close();

            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
                   "Tunel.Proceso",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }

        private void RellenaInventario(Dictionary<string, List<EndesaEntity.Tunel>> dic)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int total_registros = 0;

            int total_registros_progress_bar = 0;
            int progreso = 0;
            double percent = 0;
            forms.FrmProgressBar pb = new forms.FrmProgressBar();

            try
            {

                foreach (KeyValuePair<string, List<EndesaEntity.Tunel>> p in dic)
                {
                    for (int i = 0; i < p.Value.Count(); i++)
                        total_registros_progress_bar++;
                }

                pb.Text = "Analizando inventario";                
                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = total_registros_progress_bar + 2;

                progreso++;
                percent = (progreso / Convert.ToDouble(total_registros_progress_bar)) * 100;
                pb.progressBar.Increment(1);
                pb.progressBar.Value = progreso;
                pb.txtDescripcion.Text = "Volcando inventario anterior a histórico";
                pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                pb.Refresh();


                strSql = "replace into tunel_inventario_hist select * from tunel_inventario";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                progreso++;
                percent = (progreso / Convert.ToDouble(total_registros_progress_bar)) * 100;
                pb.progressBar.Increment(1);
                pb.txtDescripcion.Text = "Borrando tabla inventario";
                pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                pb.Refresh();

                strSql = "delete from tunel_inventario";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                foreach (KeyValuePair<string, List<EndesaEntity.Tunel>> p in dic)
                {
                    total_registros++;                    

                    for (int i = 0; i < p.Value.Count(); i++)
                    {
                        progreso++;
                        percent = (progreso / Convert.ToDouble(total_registros_progress_bar)) * 100;                        
                        pb.progressBar.Value = total_registros_progress_bar;
                        
                        pb.txtDescripcion.Text = "Guardando inventario";
                        pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                        pb.Refresh();

                        if (firstOnly)
                        {
                            sb = new StringBuilder();
                            sb.Append("replace into tunel_inventario");
                            sb.Append("(cups20, fecha_desde, fecha_hasta, usuario)");
                            sb.Append("values ");
                            sb.Append("('").Append(p.Value[i].cups20).Append("',");
                            sb.Append("'").Append(p.Value[i].startDate.ToString("yyyy-MM-dd")).Append("',");
                            sb.Append("'").Append(p.Value[i].endDate.ToString("yyyy-MM-dd")).Append("',");
                            sb.Append("'").Append(System.Environment.UserName).Append("')");

                            firstOnly = false;
                        }
                        else
                        {
                            sb.Append(",('").Append(p.Value[i].cups20).Append("',");
                            sb.Append("'").Append(p.Value[i].startDate.ToString("yyyy-MM-dd")).Append("',");
                            sb.Append("'").Append(p.Value[i].endDate.ToString("yyyy-MM-dd")).Append("',");
                            sb.Append("'").Append(System.Environment.UserName).Append("')");
                        }
                        
                    }

                    if (total_registros == 250)
                    {
                        db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString(), db.con); ;
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        total_registros = 0;
                        firstOnly = true;

                    }

                }


                if (total_registros > 0)
                {
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString(), db.con); ;
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    total_registros = 0;
                    firstOnly = true;

                }
                pb.Close();
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
                  "Tunel.RellenaInventario",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        private void GeneraExcelResultados(string fichero)
        {
            int f = 0;
            int c = 0;            

            FileInfo fileInfo = new FileInfo(fichero);
            string nombre_fichero = fileInfo.Name;
            nombre_fichero = nombre_fichero.Replace(".xlsx", "") + "_resultado_" 
                + DateTime.Now.ToString("yyyyMMdd_HHmmss") +".xlsx";

            FileInfo fileInfo2 = new FileInfo(p.GetValue("salida_excels", DateTime.Now, DateTime.Now) + nombre_fichero);
            if (fileInfo2.Exists)
                fileInfo2.Delete();

            ExcelPackage excelPackage = new ExcelPackage(fileInfo2);

            foreach (KeyValuePair<string, List<EndesaEntity.Tunel>> p in dic)
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add(p.Key);
                var headerCells = workSheet.Cells[1, 1, 1, 13];
                var headerFont = headerCells.Style.Font;
                f = 1;
                c = 1;
                headerFont.Bold = true;

                workSheet.Cells[f, c].Value = "id"; c++;
                workSheet.Cells[f, c].Value = "name"; c++;
                workSheet.Cells[f, c].Value = "versionStartDate"; c++;
                workSheet.Cells[f, c].Value = "versionEndDate"; c++;
                workSheet.Cells[f, c].Value = "volume"; c++;
                workSheet.Cells[f, c].Value = "lowerBand"; c++;
                workSheet.Cells[f, c].Value = "higherBand"; c++;
                workSheet.Cells[f, c].Value = "startDate"; c++;
                workSheet.Cells[f, c].Value = "endDate"; c++;
                workSheet.Cells[f, c].Value = "cups20"; c++;
                workSheet.Cells[f, c].Value = "cups13"; c++;
                workSheet.Cells[f, c].Value = "tarifa"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO TOTAL"; 



                foreach (EndesaEntity.Tunel pp in p.Value)
                {
                    f++;
                    c = 1;

                    workSheet.Cells[f, c].Value = pp.id; c++;
                    workSheet.Cells[f, c].Value = pp.name; c++;

                    if (pp.versionStart > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = pp.versionStart;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;                   

                    if (pp.versionEnd > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = pp.versionEnd;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    workSheet.Cells[f, c].Value = pp.volume; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = pp.lowerBand; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = pp.higherBand; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    if (pp.startDate > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = pp.startDate;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (pp.endDate > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = pp.endDate;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    workSheet.Cells[f, c].Value = pp.cups20; c++;
                    workSheet.Cells[f, c].Value = pp.cups13; c++;
                    workSheet.Cells[f, c].Value = pp.tarifa; c++;
                    workSheet.Cells[f, c].Value = pp.total_energia; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; 

                }

                var allCells = workSheet.Cells[1, 1, f, 13];
                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:M1"].AutoFilter = true;
                allCells.AutoFitColumns();
            }

            excelPackage.Save();

        }

    }
}
