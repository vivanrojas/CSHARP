using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.medida
{
    public class Kee_Reporte_Extraccion_CH
    {
        Dictionary<string, List<EndesaEntity.medida.CurvaDeCarga>> dic;

        public double total_activa { get; set; }
        public double total_reactiva { get; set; }
        public bool existe_curva { get; set; }


        utilidades.Fechas utilFecha = new utilidades.Fechas();
        
        public Kee_Reporte_Extraccion_CH()
        {
            Console.WriteLine("Cargando datos de kee_reporte_extraccion_ch");
            dic = Load();
            Guardar_Kee_Reporte_Extraccion_CH_Horizontal(dic);
        }

        public Kee_Reporte_Extraccion_CH(bool solo_informe)
        {
            
        }


        private Dictionary<string, List<EndesaEntity.medida.CurvaDeCarga>> Load()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string cups22 = "";
            DateTime fecha = new DateTime();
            int hora = 0;
            EndesaEntity.medida.CurvaDeCarga c;
            bool encontrado = false;
            string fuente = "";

            Dictionary<string, List<EndesaEntity.medida.CurvaDeCarga>> d = new Dictionary<string, List<EndesaEntity.medida.CurvaDeCarga>>();

            try
            {
                strSql = "SELECT cups22, fuente, periodo, fecha, hora, ae, `_as`, r1, r2, r3, r4,"
                    + " cal_ae, cal_as, cal_r1, cal_r2, cal_r3, cal_r4"
                    + " FROM med.kee_reporte_extraccion_ch";
                    // + " where cups22 = 'ES0021000003296660KK2R' and fecha = '2020-02-13'";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha = Convert.ToDateTime(r["fecha"]);
                    hora = Convert.ToInt32(r["hora"].ToString().Substring(0, 2));
                    cups22 = r["cups22"].ToString();
                    fuente = r["fuente"].ToString();

                    if (hora == 0)
                    {
                        fecha = fecha.AddDays(-1);
                        if (fecha.Month != 3 && fecha.Month != 10)
                        {
                            hora = 24;
                        }
                        else if (fecha == utilFecha.UltimoDomingoMarzo(fecha))
                        {
                            hora = 23;
                        }
                        else if (fecha == utilFecha.UltimoDomingoOctubre(fecha))
                        {
                            hora = 25;
                        }
                        else
                            hora = 24;
                        
                    }
                    List<EndesaEntity.medida.CurvaDeCarga> o;
                    if(!d.TryGetValue(cups22,out o))
                    {
                        o = new List<EndesaEntity.medida.CurvaDeCarga>();
                        c = new EndesaEntity.medida.CurvaDeCarga();
                        c.fecha = fecha;
                        c.fuente = r["fuente"].ToString();

                        if(r["ae"] != System.DBNull.Value)
                        {
                            c.numPeriodosHorarios = c.numPeriodosHorarios + 1;
                            c.horaria_activa[hora - 1] = Convert.ToDouble(r["ae"]);
                            c.existe_horaria_activa[hora - 1] = true;
                            c.total_activa = c.total_activa + c.horaria_activa[hora - 1];
                        }
                            
                        if (r["r1"] != System.DBNull.Value)
                        {
                            c.horaria_reactiva[hora - 1] = Convert.ToDouble(r["r1"]);
                            c.existe_horaria_reactiva[hora - 1] = true;
                            c.total_reactiva = c.total_reactiva + c.horaria_reactiva[hora - 1];
                        }
                        
                        
                        o.Add(c);
                        d.Add(cups22, o);
                    }
                    else
                    {
                        encontrado = false;
                        // c = o.Find(z => z.fecha == fecha);
                        foreach(EndesaEntity.medida.CurvaDeCarga p in o)
                        {
                            if(p.fecha == fecha && p.fuente == fuente)
                            {
                                encontrado = true;
                                if (r["ae"] != System.DBNull.Value)
                                {
                                    p.numPeriodosHorarios = p.numPeriodosHorarios + 1;
                                    p.horaria_activa[hora - 1] = Convert.ToDouble(r["ae"]);
                                    p.existe_horaria_activa[hora - 1] = true;
                                    p.total_activa = p.total_activa + p.horaria_activa[hora - 1];
                                }
                                    
                                if (r["r1"] != System.DBNull.Value)
                                {
                                    p.horaria_reactiva[hora - 1] = Convert.ToDouble(r["r1"]);
                                    p.existe_horaria_reactiva[hora - 1] = true;
                                    p.total_reactiva = p.total_reactiva + p.horaria_reactiva[hora - 1];
                                }
                                    
                            }
                            
                        }
                        if(!encontrado)
                        {
                            EndesaEntity.medida.CurvaDeCarga cc = new EndesaEntity.medida.CurvaDeCarga();
                            cc.fecha = fecha;
                            cc.fuente = r["fuente"].ToString();
                            if (r["ae"] != System.DBNull.Value)
                            {
                                cc.numPeriodosHorarios = cc.numPeriodosHorarios + 1;
                                cc.horaria_activa[hora - 1] = Convert.ToDouble(r["ae"]);
                                cc.existe_horaria_activa[hora - 1] = true;
                                cc.total_activa = cc.total_activa + cc.horaria_activa[hora - 1];
                            }                               
                            
                            if (r["r1"] != System.DBNull.Value)
                            {
                                cc.horaria_reactiva[hora - 1] = Convert.ToDouble(r["r1"]);
                                cc.existe_horaria_reactiva[hora - 1] = true;
                                cc.total_reactiva = cc.total_reactiva + cc.horaria_reactiva[hora - 1];
                            }
                            o.Add(cc);
                        }
                    }

                    

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        public void GetCurva(string cups22, DateTime fd, DateTime fh, string fuente)
        {
                        
            List<EndesaEntity.medida.CurvaDeCarga> o;
            existe_curva = false;
            total_activa = 0;
            total_reactiva = 0;
            if (dic.TryGetValue(cups22, out o))
            {
                
                foreach (EndesaEntity.medida.CurvaDeCarga p in o)
                {
                    
                    if (p.fuente == fuente && (p.fecha >= fd && p.fecha <= fh))
                    {
                        existe_curva = true;
                        total_activa = total_activa + p.total_activa;
                        total_reactiva = total_reactiva + p.total_reactiva;
                    }
                    
                }
                
            }

            
        }

        private void Guardar_Kee_Reporte_Extraccion_CH_Horizontal(Dictionary<string, List<EndesaEntity.medida.CurvaDeCarga>> d)
        {
            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int registros = 0;
            int totalregistros = 0;
            try
            {
                foreach (KeyValuePair<string, List<EndesaEntity.medida.CurvaDeCarga>> p in d)
                {                    

                    foreach(EndesaEntity.medida.CurvaDeCarga j in p.Value)
                    {
                        registros++;
                        totalregistros++;

                        if (firstOnly)
                        {
                            sb.Append("replace into kee_reporte_extraccion_ch_horizontal");
                            sb.Append(" (cups20, cups22, fecha, fuente, total_a, total_r");
                            for (int i = 1; i <= 25; i++)
                                sb.Append(",a" + i);

                            for (int i = 1; i <= 25; i++)
                                sb.Append(",r" + i);

                            sb.Append(") values ");
                            firstOnly = false;
                        }

                        sb.Append("('").Append(p.Key.Substring(0, 20)).Append("',");
                        sb.Append("'").Append(p.Key).Append("',");
                        sb.Append("'").Append(j.fecha.ToString("yyyy-MM-dd")).Append("',");
                        sb.Append("'").Append(j.fuente).Append("',");
                        sb.Append(j.total_activa.ToString().Replace(",", ".")).Append(",");
                        sb.Append(j.total_reactiva.ToString().Replace(",", "."));

                        for (int i = 0; i < 25; i++)
                        {
                            if (j.existe_horaria_activa[i])
                                sb.Append(",").Append(j.horaria_activa[i].ToString().Replace(",", "."));
                            else
                                sb.Append(",null");
                        }
                            

                        for (int i = 0; i < 25; i++)
                        {
                            if (j.existe_horaria_reactiva[i])
                                sb.Append(",").Append(j.horaria_reactiva[i].ToString().Replace(",", "."));
                            else
                                sb.Append(",null");
                        }
                            

                        sb.Append("),");

                        if (registros == 250)
                        {
                            Console.CursorLeft = 0;
                            Console.Write("Guardando " + totalregistros);
                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.MED);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            registros = 0;
                        }
                    }                       
                    
                }

                if (registros > 0)
                {
                    Console.CursorLeft = 0;
                    Console.Write("Guardando " + totalregistros);
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    registros = 0;
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }


        private List<EndesaEntity.medida.Kee_Extraccion_Formulas> Carga(List<EndesaEntity.medida.Kee_Extraccion_Formulas> lista)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";            
            double percent = 0;
            int j = 0;
                      

            forms.FrmProgressBar pb = new forms.FrmProgressBar();
            try
            {
                pb.Text = "Consultando datos ...";
                pb.Show();
                pb.progressBar.Minimum = 0;                
                pb.progressBar.Maximum = 100;

                foreach (EndesaEntity.medida.Kee_Extraccion_Formulas p in lista)
                {
                    j++;
                    percent = (j / Convert.ToDouble(lista.Count())) * 100;
                    pb.progressBar.Value = Convert.ToInt32(Math.Round(percent,0));
                    pb.txtDescripcion.Text = "Consultando " 
                        + p.cups22 + " --> " + p.tipo + " --> " + p.fuente;
                        
                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    pb.Refresh();

                    if (p.fuente == "CH")
                    {
                        strSql = "SELECT cups20, cups22, fuente, periodo, fecha, hora, estacion,"
                            + " ae, c._as, r1, r2, r3, r4, cal_ae, cal_as,"
                            + " cal_r1, cal_r2, cal_r3, cal_r4,"
                            + " metodo_obtencion, indicador_firmeza"
                            + " FROM kee_reporte_extraccion_ch c where"
                            + " c.cups22 = '" + p.cups22 + "' and"
                            + " (fecha >= '" + p.fecha_desde.ToString("yyyy-MM-dd") + "' and"
                            + "  fecha <= '" + p.fecha_hasta.ToString("yyyy-MM-dd") + "') and"
                            + " fuente = '" + p.tipo + "'";

                    }
                    else
                    {
                        strSql = "SELECT cups20, cups22, fuente, periodo, fecha, hora, estacion,"
                            + " ae, c._as, r1, r2, r3, r4, cal_ae, cal_as,"
                            + " cal_r1, cal_r2, cal_r3, cal_r4,"
                            + " metodo_obtencion, indicador_firmeza"
                            + " FROM kee_reporte_extraccion_cch c where"
                            + " c.cups22 = '" + p.cups22 + "' and"
                            + " (fecha >= '" + p.fecha_desde.ToString("yyyy-MM-dd") + "' and"
                            + "  fecha <= '" + p.fecha_hasta.ToString("yyyy-MM-dd") + "') and"
                            + " fuente = '" + p.tipo + "'";

                    }

                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        EndesaEntity.medida.Kee_Reporte_Extraccion c =
                                new EndesaEntity.medida.Kee_Reporte_Extraccion();

                        if (r["cups20"] != System.DBNull.Value)
                            c.cups20 = r["cups20"].ToString();
                        if (r["cups22"] != System.DBNull.Value)
                            c.cups22 = r["cups22"].ToString();
                        if (r["fuente"] != System.DBNull.Value)
                            c.fuente = r["fuente"].ToString();
                        if (r["fecha"] != System.DBNull.Value)
                            c.fecha = Convert.ToDateTime(r["fecha"]);
                        if (r["hora"] != System.DBNull.Value)
                            c.hora = r["hora"].ToString();
                        if (r["estacion"] != System.DBNull.Value)
                            c.estacion = Convert.ToInt32(r["estacion"]);

                        if (r["ae"] != System.DBNull.Value)
                            c.energia[0] = Convert.ToDouble(r["ae"]);
                        if (r["_as"] != System.DBNull.Value)
                            c.energia[1] = Convert.ToDouble(r["_as"]);
                        for (int i = 1; i <= 4; i++)
                            if (r["r" + i] != System.DBNull.Value)
                                c.energia[i + 1] = Convert.ToDouble(r["r" + i]);

                        if (r["cal_ae"] != System.DBNull.Value)
                            c.cal[0] = Convert.ToInt32(r["cal_ae"]);
                        if (r["cal_as"] != System.DBNull.Value)
                            c.cal[1] = Convert.ToInt32(r["cal_as"]);
                        for (int i = 1; i <= 4; i++)
                            if (r["cal_r" + i] != System.DBNull.Value)
                                c.cal[i + 1] = Convert.ToInt32(r["cal_r" + i]);

                        if (r["metodo_obtencion"] != System.DBNull.Value)
                            c.metodo_obtencion = r["metodo_obtencion"].ToString();

                        if (r["indicador_firmeza"] != System.DBNull.Value)
                            c.indicador_firmeza = r["indicador_firmeza"].ToString();

                        p.lista.Add(c);

                    }
                    db.CloseConnection();

                }
                pb.progressBar.Value = 100;
                pb.Close();

                return lista;
            }
            catch(Exception e)
            {
                pb.Close();
                return null;
            }
        }

        public void InformeExcel(List<EndesaEntity.medida.Kee_Extraccion_Formulas> lista)
        {
            int c = 1;
            int f = 1;
            SaveFileDialog save;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage;
            int num_hojas = 0;
            int total_filas = 0;
            int j = 0;
            double percent = 0;
            forms.FrmProgressBar pb = new forms.FrmProgressBar();

            try
            {

                List<EndesaEntity.medida.Kee_Extraccion_Formulas> d =
                    Carga(lista);

                if (d != null &&  d.Count > 0)
                {

                    save = new SaveFileDialog();
                    save.Title = "Ubicación del informe";
                    save.AddExtension = true;
                    save.DefaultExt = "xlsx";
                    DialogResult result = save.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        Cursor.Current = Cursors.WaitCursor;
                        
                        pb.Text = "Generando Excel ...";
                        pb.Show();
                        pb.progressBar.Minimum = 0;
                        pb.progressBar.Maximum = 100;

                        foreach (EndesaEntity.medida.Kee_Extraccion_Formulas p in d)
                            total_filas += p.lista.Count();

                        FileInfo fileInfo = new FileInfo(save.FileName);
                        if (fileInfo.Exists)
                            fileInfo.Delete();

                        num_hojas++;
                        excelPackage = new ExcelPackage(fileInfo);
                        var workSheet = excelPackage.Workbook.Worksheets.Add("CURVAS");
                        var headerCells = workSheet.Cells[1, 1, 1, 23];
                        var headerFont = headerCells.Style.Font;
                        var allCells = workSheet.Cells[1, 1, f, 23];

                        workSheet.View.FreezePanes(2, 1);
                        headerFont.Bold = true;

                        #region Cabecera
                        workSheet.Cells[f, c].Value = "CUPS SOLICITADO";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        workSheet.Cells[f, c].Value = "FUENTE SOLICITADA";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        workSheet.Cells[f, c].Value = "TIPO SOLICITADO";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        workSheet.Cells[f, c].Value = "CUPS20";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        workSheet.Cells[f, c].Value = "CUPS22";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        workSheet.Cells[f, c].Value = "FUENTE";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        workSheet.Cells[f, c].Value = "FECHA";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        workSheet.Cells[f, c].Value = "HORA";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        workSheet.Cells[f, c].Value = "ESTACION";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        workSheet.Cells[f, c].Value = "AE";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        workSheet.Cells[f, c].Value = "AS";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        for (int i = 1; i <= 4; i++)
                        {
                            workSheet.Cells[f, c].Value = "R" + i;
                            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            c++;
                        }

                        workSheet.Cells[f, c].Value = "CAL_AE";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        workSheet.Cells[f, c].Value = "CAL_AS";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        for (int i = 1; i <= 4; i++)
                        {
                            workSheet.Cells[f, c].Value = "CAL_R" + i;
                            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            c++;
                        }

                        workSheet.Cells[f, c].Value = "MÉTODO OBTENCIÓN";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        workSheet.Cells[f, c].Value = "INDICADOR FIRMEZA";
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        #endregion

                        foreach (EndesaEntity.medida.Kee_Extraccion_Formulas p in d)
                        {


                            

                            if (f + p.lista.Count() > 1000000)
                            {

                                allCells = workSheet.Cells[1, 1, f, 23];
                                workSheet.Cells["A1:U1"].AutoFilter = true;

                                headerFont.Bold = true;
                                //allCells.AutoFitColumns();

                                num_hojas++;
                                f = 1;
                                c = 1;
                                workSheet = excelPackage.Workbook.Worksheets.Add("CURVAS_" + num_hojas);
                                headerCells = workSheet.Cells[1, 1, 1, 46];
                                headerFont = headerCells.Style.Font;

                                workSheet.View.FreezePanes(2, 1);
                                headerFont.Bold = true;

                                #region Cabecera
                                workSheet.Cells[f, c].Value = "CUPS SOLICITADO";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                workSheet.Cells[f, c].Value = "FUENTE SOLICITADA";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                workSheet.Cells[f, c].Value = "TIPO SOLICITADA";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                workSheet.Cells[f, c].Value = "CUPS20";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                workSheet.Cells[f, c].Value = "CUPS22";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                workSheet.Cells[f, c].Value = "FUENTE";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                workSheet.Cells[f, c].Value = "FECHA";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                workSheet.Cells[f, c].Value = "HORA";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                workSheet.Cells[f, c].Value = "ESTACION";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                workSheet.Cells[f, c].Value = "AE";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                workSheet.Cells[f, c].Value = "AS";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                for (int i = 1; i <= 4; i++)
                                {
                                    workSheet.Cells[f, c].Value = "R" + i;
                                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                    c++;
                                }

                                workSheet.Cells[f, c].Value = "CAL_AE";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                workSheet.Cells[f, c].Value = "CAL_AS";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                for (int i = 1; i <= 4; i++)
                                {
                                    workSheet.Cells[f, c].Value = "CAL_R" + i;
                                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                    c++;
                                }

                                workSheet.Cells[f, c].Value = "MÉTODO OBTENCIÓN";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                workSheet.Cells[f, c].Value = "INDICADOR FIRMEZA";
                                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                c++;

                                #endregion
                            }

                            if (p.lista.Count == 0)
                            {
                                f++;
                                c = 1;
                                workSheet.Cells[f, c].Value = p.cups20; c++;
                                workSheet.Cells[f, c].Value = p.fuente; c++;
                                workSheet.Cells[f, c].Value = p.tipo; c++;
                            }
                            else
                                foreach (EndesaEntity.medida.Kee_Reporte_Extraccion pp in p.lista)
                                {

                                    j++;
                                    percent = (j / Convert.ToDouble(total_filas)) * 100;
                                    pb.progressBar.Value = Convert.ToInt32(Math.Round(percent, 0));
                                    pb.txtDescripcion.Text = "Exportando... "
                                        + p.cups20 + " --> " + p.tipo + " --> " + p.fuente;
                                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                                    pb.Refresh();


                                    f++;
                                    c = 1;

                                    if (f == 3 && num_hojas == 1)
                                    {
                                        allCells = workSheet.Cells[1, 1, f, 23];
                                        allCells.AutoFitColumns();
                                    }

                                    workSheet.Cells[f, c].Value = p.cups20; c++;
                                    workSheet.Cells[f, c].Value = p.fuente; c++;
                                    workSheet.Cells[f, c].Value = p.tipo; c++;
                                    workSheet.Cells[f, c].Value = pp.cups20; c++;
                                    workSheet.Cells[f, c].Value = pp.cups22; c++;
                                    workSheet.Cells[f, c].Value = pp.fuente; c++;

                                    if (pp.fecha > DateTime.MinValue)
                                    {
                                        workSheet.Cells[f, c].Value = pp.fecha;
                                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    }
                                    c++;

                                    if (pp.fecha > DateTime.MinValue)
                                    {
                                        workSheet.Cells[f, c].Value = pp.hora;
                                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortTimePattern;
                                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    }
                                    c++;

                                    workSheet.Cells[f, c].Value = pp.estacion; c++;

                                    for (int i = 0; i < 6; i++)
                                    {
                                        workSheet.Cells[f, c].Value = pp.energia[i];
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        c++;
                                    }

                                    for (int i = 0; i < 6; i++)
                                    {
                                        workSheet.Cells[f, c].Value = pp.cal[i];
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        c++;
                                    }
                                    workSheet.Cells[f, c].Value = pp.metodo_obtencion; c++;
                                    workSheet.Cells[f, c].Value = pp.indicador_firmeza; c++;

                                }
                        }

                        allCells = workSheet.Cells[1, 1, f, 23];
                        workSheet.Cells["A1:W1"].AutoFilter = true;


                        headerFont.Bold = true;

                        excelPackage.Save();

                        Cursor.Current = Cursors.Default;
                        pb.Close();

                        MessageBox.Show("Informe terminado.",
                               "Informe Excel",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Information);

                        DialogResult result3 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                        if (result3 == DialogResult.Yes)
                            System.Diagnostics.Process.Start(save.FileName);


                    }
                }

            }
            catch (Exception e)
            {
                pb.Close();
                MessageBox.Show("Error" + e.Message,
                               "Informe Excel",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
            }
        }

        public void InformePO1011(List<EndesaEntity.medida.Kee_Extraccion_Formulas> lista)
        {


            int c = 1;
            int f = 1;
            SaveFileDialog save;
            string linea = "";
            int total_filas = 0;
            int j = 0;
            double percent = 0;
            forms.FrmProgressBar pb = new forms.FrmProgressBar();

            try
            {
                List<EndesaEntity.medida.Kee_Extraccion_Formulas> d =
                   Carga(lista);

                if (d != null && d.Count > 0)
                {
                    save = new SaveFileDialog();
                    save.Title = "Ubicación del informe";
                    save.AddExtension = true;
                    save.DefaultExt = "txt";
                    DialogResult result = save.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        Cursor.Current = Cursors.WaitCursor;

                        pb.Text = "Generando archivo ...";
                        pb.Show();
                        pb.progressBar.Minimum = 0;
                        pb.progressBar.Maximum = 100;

                        foreach (EndesaEntity.medida.Kee_Extraccion_Formulas p in d)
                            total_filas += p.lista.Count();

                        FileInfo fileInfo = new FileInfo(save.FileName);
                        if (fileInfo.Exists)
                            fileInfo.Delete();

                        StreamWriter swa = new StreamWriter(fileInfo.FullName, false);

                        foreach (EndesaEntity.medida.Kee_Extraccion_Formulas p in d)
                            foreach (EndesaEntity.medida.Kee_Reporte_Extraccion pp in p.lista)
                            {
                                j++;
                                percent = (j / Convert.ToDouble(total_filas)) * 100;
                                pb.progressBar.Value = Convert.ToInt32(Math.Round(percent, 0));
                                pb.txtDescripcion.Text = "Exportando... "
                                    + p.cups22 + " --> " + p.tipo + " --> " + p.fuente;
                                pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                                pb.Refresh();

                                linea = pp.cups22 + ";"
                                    + "11" + ";"
                                    + pp.fecha.ToString("yyyy/MM/dd") + " "
                                    + Convert.ToDateTime(pp.hora).AddHours(1).ToString("HH:mm:ss") + ";"
                                    + pp.estacion;

                                for(int w = 0; w <= 5; w++)
                                {
                                    linea += ";" + pp.energia[w] + ";" // AE
                                    + pp.cal[w];  
                                    
                                }

                                    
                                linea += ";0;128;0;128;1;0";

                                

                                swa.WriteLine(linea);
                            }

                        swa.Close();
                        Cursor.Current = Cursors.Default;
                        pb.Close();

                        MessageBox.Show("Informe terminado.",
                              "Informe",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Information);
                    }
                }
            }
            catch(Exception e)
            {
                pb.Close();
                MessageBox.Show("Error" + e.Message,
                               "Informe PO1011",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
            }
        }

        private void PintaCabecera()
        {

        }

    }
}

