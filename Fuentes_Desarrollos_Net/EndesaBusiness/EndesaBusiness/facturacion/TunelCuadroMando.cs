using EndesaBusiness.servidores;
using Microsoft.Graph;
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

namespace EndesaBusiness.facturacion
{
    public class TunelCuadroMando: EndesaEntity.TunelContrato
    {
        Dictionary<string, List<EndesaEntity.Tunel>> dic;
        public Dictionary<string,EndesaEntity.TunelContrato> dic_contratos { get; set; }
        
        EndesaBusiness.utilidades.Param p;
        public TunelCuadroMando(DateTime fd, DateTime fh, string cliente)
        {
            p = new EndesaBusiness.utilidades.Param("tunel_param", MySQLDB.Esquemas.FAC);
            dic_contratos = CargaContratos(fd, fh, cliente);
        }

        public void CargaExcel(string fichero, bool mensual)
        {
            dic = ProcesaExcel(fichero, mensual);
            ActualizaContratos(dic);

           
        }

        private Dictionary<string, EndesaEntity.TunelContrato> CargaContratos(DateTime fd, DateTime fh, string cliente)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            EndesaEntity.TunelContrato o;

            Dictionary<string, EndesaEntity.TunelContrato> d =
                new Dictionary<string, EndesaEntity.TunelContrato>();

            EndesaBusiness.contratacion.PS_AT psat;
            EndesaBusiness.medida.Pendiente pendiente_SCE;
            string clave = "";


            try
            {

                psat = new contratacion.PS_AT();
                pendiente_SCE = new medida.Pendiente();
                pendiente_SCE.CargaPendiente();

                strSql = "SELECT cliente, fecha_inicio_tunel, fecha_final_tunel, consumo_referencia,"
                    + " banda_inferior_pct, banda_superior_pct, banda_inferior_gwh, banda_superior_gwh,"
                    + " consumo_real_kwh, consumo_real_gwh, aplica_tunel, formula_antigua, num_cups,"
                    + " num_cups_completos, fecha_minima_periodo, fecha_maxima_periodo"
                    + " FROM tunel_contratos_resumen where"
                    + " (fecha_inicio_tunel <= '" + fh.ToString("yyyy-MM-dd") + "' AND"
                     + " fecha_final_tunel >= '" + fd.ToString("yyyy-MM-dd") + "')";

                if (cliente != null || cliente != "")
                    strSql += " AND cliente like '%" + cliente + "%'"; 

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while(r.Read())
                {
                    EndesaEntity.TunelContrato c = new EndesaEntity.TunelContrato();

                    if (r["cliente"] != System.DBNull.Value)
                        c.cliente = r["cliente"].ToString();

                    if (r["fecha_inicio_tunel"] != System.DBNull.Value)
                        c.fecha_inicio_tunel = Convert.ToDateTime(r["fecha_inicio_tunel"]);

                    if (r["fecha_final_tunel"] != System.DBNull.Value)
                        c.fecha_final_tunel  = Convert.ToDateTime(r["fecha_final_tunel"]);

                    if (r["consumo_referencia"] != System.DBNull.Value)
                        c.consumo_referencia = Convert.ToDouble(r["consumo_referencia"]);

                    if (r["banda_inferior_pct"] != System.DBNull.Value)
                        c.banda_inferior_pct = Convert.ToDouble(r["banda_inferior_pct"]);

                    if (r["banda_superior_pct"] != System.DBNull.Value)
                        c.banda_superior_pct = Convert.ToDouble(r["banda_superior_pct"]);

                    if (r["banda_inferior_gwh"] != System.DBNull.Value)
                        c.banda_inferior_gwh = Convert.ToDouble(r["banda_inferior_gwh"]);

                    if (r["banda_superior_gwh"] != System.DBNull.Value)
                        c.banda_superior_gwh = Convert.ToDouble(r["banda_superior_gwh"]);

                    if (r["consumo_real_kwh"] != System.DBNull.Value) 
                        c.consumo_real_kwh = Convert.ToDouble(r["consumo_real_kwh"]);

                    if (r["consumo_real_gwh"] != System.DBNull.Value)
                        c.consumo_real_gwh = Convert.ToDouble(r["consumo_real_gwh"]);
                    
                    c.aplica_tunel = r["aplica_tunel"].ToString() == "S";                    
                    c.formula_antigua = r["formula_antigua"].ToString() == "S";

                    if (r["num_cups"] != System.DBNull.Value)
                        c.total_cups = Convert.ToInt32(r["num_cups"]);

                    if (r["num_cups_completos"] != System.DBNull.Value)
                        c.total_incompletos = Convert.ToInt32(r["num_cups_completos"]);


                    clave = c.cliente + c.fecha_inicio_tunel.ToString("yyyyMMdd") + c.fecha_final_tunel.ToString("yyyyMMdd");

                    if (!d.TryGetValue(clave, out o))
                        d.Add(clave, c);


                }
                db.CloseConnection();



                strSql = "SELECT name, tc.version_start, tc.version_end, tc.volume, tc.lowerband, tc.higherband,"
                     + " tc.start_date, tc.end_date, tc.cups20, tc.fecha_primera_factura, tc.fecha_ultima_factura,"
                     + " tc.tarifa, tc.consumo_total"
                     + " FROM fact.tunel_contratos tc where"
                     + " (tc.start_date <= '" + fh.ToString("yyyy-MM-dd") + "' AND"
                     + " tc.end_date >= '" + fd.ToString("yyyy-MM-dd") + "') ";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {


                    EndesaEntity.Tunel c = new EndesaEntity.Tunel();
                    c.name = r["name"].ToString();
                    c.versionStart = Convert.ToDateTime(r["version_start"]);
                    c.versionEnd = Convert.ToDateTime(r["version_end"]);
                    c.startDate = Convert.ToDateTime(r["start_date"]);
                    c.endDate = Convert.ToDateTime(r["end_date"]);
                    c.volume = Convert.ToDouble(r["volume"]);
                    c.lowerBand = Convert.ToDouble(r["lowerband"]);
                    c.higherBand = Convert.ToDouble(r["higherband"]);
                    c.cups20 = r["cups20"].ToString();

                    if (r["tarifa"] != System.DBNull.Value)
                        c.tarifa = r["tarifa"].ToString();

                    if (r["fecha_primera_factura"] != System.DBNull.Value)
                        c.fecha_primera_factura = Convert.ToDateTime(r["fecha_primera_factura"]);

                    if (r["fecha_ultima_factura"] != System.DBNull.Value)
                        c.fecha_ultima_factura = Convert.ToDateTime(r["fecha_ultima_factura"]);

                    if(c.fecha_ultima_factura != c.endDate)
                    {
                        if (c.tarifa != "2.0TD" && c.tarifa != "3.0TD")
                            c.es_baja = !psat.ExisteAlta(c.cups20);
                        else
                            c.comentario = "Punto de baja tensión";

                        // si no es baja y no tiene la facturación completa
                        // comprobamos LTP´s
                        if (!c.es_baja)
                        {
                            pendiente_SCE.GetCUPS20(c.cups20);
                            if (pendiente_SCE.existe)
                            {
                                if(pendiente_SCE.aaaammPdte <= Convert.ToInt32(c.endDate.ToString("yyyyMM")))
                                {
                                    c.comentario = "Tiene LTP --> " + pendiente_SCE.subsEstado;
                                    c.tiene_ltp = true;
                                }
                                
                            }
                            

                        }

                    }


                    if (r["consumo_total"] != System.DBNull.Value)
                        c.total_energia = Convert.ToDouble(r["consumo_total"]) / 1000000;




                    clave = c.name + c.startDate.ToString("yyyyMMdd") + c.endDate.ToString("yyyyMMdd");

                    if (d.TryGetValue(clave, out o))
                    {
                        if (c.tarifa == "2.0TD" || c.tarifa == "3.0TD")
                        {
                            o.comentario = "El contrato tiene puntos de baja tensión";
                            o.baja_tension = true;
                        }                            

                        o.ltps = c.tiene_ltp;

                        o.lista.Add(c);
                    }


                }
                db.CloseConnection();
                return d;
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
                  "TunelCuadroMando - CargaContratos",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Warning);
                return null;
            }
        }

        
        private void ActualizaContratos(Dictionary<string, List<EndesaEntity.Tunel>> dic)
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
            //forms.FrmProgressBar pb = new forms.FrmProgressBar();

            try
            {

                strSql = "DELETE FROM tunel_contratos";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con); ;
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "DELETE FROM tunel_contratos_resumen";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con); ;
                command.ExecuteNonQuery();
                db.CloseConnection();


                //foreach (KeyValuePair<string, List<EndesaEntity.Tunel>> p in dic)
                //{
                //    for (int i = 0; i < p.Value.Count(); i++)
                //        total_registros_progress_bar++;
                //}

                //pb.Text = "Analizando contratos";
                //pb.Show();
                //pb.progressBar.Step = 1;
                //pb.progressBar.Maximum = total_registros_progress_bar + 2;

                progreso++;                

                foreach (KeyValuePair<string, List<EndesaEntity.Tunel>> p in dic)
                {
                    

                    for (int i = 0; i < p.Value.Count(); i++)
                    {
                        progreso++;
                        //percent = (progreso / Convert.ToDouble(total_registros_progress_bar)) * 100;
                        //pb.progressBar.Value = total_registros_progress_bar;

                        //pb.txtDescripcion.Text = "Guardando contratos";
                        //pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                        //pb.Refresh();

                        total_registros++;

                        if (firstOnly)
                        {
                            sb = new StringBuilder();
                            sb.Append("replace into tunel_contratos");
                            sb.Append(" (name, version_start, version_end,");
                            sb.Append(" volume, lowerband, higherband, start_date,");
                            sb.Append(" end_date, cups20, tarifa, created_by, created_date)");
                            sb.Append(" values ");                          

                            firstOnly = false;
                        }

                        sb.Append("('").Append(p.Value[i].name).Append("',");
                        sb.Append("'").Append(p.Value[i].versionStart.ToString("yyyy-MM-dd")).Append("',");
                        sb.Append("'").Append(p.Value[i].versionEnd.ToString("yyyy-MM-dd")).Append("',");
                        sb.Append(p.Value[i].volume.ToString().Replace(",", ".")).Append(",");
                        sb.Append(p.Value[i].lowerBand.ToString().Replace(",", ".")).Append(",");
                        sb.Append(p.Value[i].higherBand.ToString().Replace(",", ".")).Append(",");
                        sb.Append("'").Append(p.Value[i].startDate.ToString("yyyy-MM-dd")).Append("',");
                        sb.Append("'").Append(p.Value[i].endDate.ToString("yyyy-MM-dd")).Append("',");
                        sb.Append("'").Append(p.Value[i].cups20).Append("',");
                        sb.Append("'").Append(p.Value[i].tarifa).Append("',");
                        sb.Append("'").Append(System.Environment.UserName).Append("',");
                        sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                        if (total_registros == 250)
                        {
                            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con); ;
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            total_registros = 0;
                            firstOnly = true;
                            sb = null;

                        }

                    }

                    

                }


                if (total_registros > 0)
                {
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con); ;
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    total_registros = 0;
                    firstOnly = true;
                    sb = null;

                }
                // pb.Close();


                ActualizaContratosConFacturas();

                strSql = "REPLACE INTO tunel_contratos_resumen"
                    + " SELECT c.name AS cliente,"
                    + " c.start_date AS fecha_inicio_tunel,"
                    + " c.end_date AS fecha_final_tunel,"
                    + " c.volume AS consumo_referencia,"
                    + " c.lowerband AS banda_inferior_pct,"
                    + " c.higherband AS banda_superior_pct,"
                    + " (c.volume * c.lowerband) / 100 AS banda_inferior_gwh,"
                    + " (c.volume * c.higherband) / 100 AS banda_superior_gwh,"
                    + " SUM(c.consumo_total) AS consumo_real_kwh,"
                    + " SUM(c.consumo_total) / 1000000 AS consumo_real_gwh,"
                    + " if("
                    + " (SUM(c.consumo_total) / 1000000 < (c.volume * c.lowerband) / 100) OR"
                    + " (SUM(c.consumo_total) / 1000000 > (c.volume * c.higherband) / 100),'S','N') AS aplica_tunel,"
                    + " 'N' as formula_antigua, count(*) as num_cups, 0 as num_cups_completos,"
                    + " min(c.start_date) as fecha_minima_periodo,"
                    + " max(c.end_date) as fecha_maxima_periodo,"
                    + " '" + System.Environment.UserName.ToUpper() + "' as created_by,"
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' as created_date,"
                    + "'null' as last_update_by,"
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' as last_update_date"
                    + " FROM tunel_contratos c GROUP BY c.name, c.start_date, c.end_date";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con); ;
                command.ExecuteNonQuery();
                db.CloseConnection();
                              


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "Tunel.RellenaInventario",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        public void ActualizaContratosConFacturas()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

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

            
            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }
        
        public void ActualizaContratosResumen()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            strSql = "UPDATE tunel_contratos_resumen r"
                    + " INNER JOIN"
                    + " (SELECT c.name AS cliente,"
                    + " c.start_date AS fecha_inicio_tunel,"
                    + " c.end_date AS fecha_final_tunel,"
                    + " c.volume AS consumo_referencia,"
                    + " c.lowerband AS banda_inferior_pct,"
                    + " c.higherband AS banda_superior_pct,"
                    + " (c.volume * c.lowerband) / 100 AS banda_inferior_gwh,"
                    + " (c.volume * c.higherband) / 100 AS banda_superior_gwh,"
                    + " SUM(c.consumo_total) AS consumo_real_kwh,"
                    + " SUM(c.consumo_total) / 1000000 AS consumo_real_gwh,"
                    + " if("
                    + " (SUM(c.consumo_total) / 1000000 < (c.volume * c.lowerband) / 100) OR"
                    + " (SUM(c.consumo_total) / 1000000 > (c.volume * c.higherband) / 100),'S','N') AS aplica_tunel,"
                    + " 'N' as formula_antigua, count(*) as num_cups, 0 as num_cups_completos,"
                    + " min(c.start_date) as fecha_minima_periodo,"
                    + " max(c.end_date) as fecha_maxima_periodo,"
                    + " '" + System.Environment.UserName.ToUpper() + "' as created_by,"
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' as created_date,"
                    + "'null' as last_update_by,"
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' as last_update_date"
                    + " FROM tunel_contratos c GROUP BY c.name, c.start_date, c.end_date) c ON"
                    + " c.cliente = r.cliente and"
                    + " c.fecha_inicio_tunel = r.fecha_inicio_tunel and"
                    + " c.fecha_final_tunel = r.fecha_final_tunel"
                    + " set r.consumo_real_kwh = c.consumo_real_kwh,"
                    + " r.consumo_real_gwh = c.consumo_real_gwh,"
                    + " r.aplica_tunel = c.aplica_tunel,"
                    + " r.last_update_by = c.last_update_by";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con); ;
            command.ExecuteNonQuery();
            db.CloseConnection();

          
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

                    if (cabecera != p.GetValue("cabecera_excel"))
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
                    for (int i = 1; i < 100000000; i++)
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
                                tt.endDate = new DateTime(mes.Year, mes.Month, dias_del_mes);

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


        public void CargaExcelConsumos(string fichero, bool mensual)
        {

            List<EndesaEntity.Tunel> o;
            int c = 1;
            int f = 1;
            int total_hojas_excel = 0;
            bool firstOnly = true;
            string cabecera = "";
            FileStream fs;
            ExcelPackage excelPackage;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            

            int anio = 0;
            int mes = 0;

            Dictionary<string, List<EndesaEntity.Tunel>> dic_consumos = new Dictionary<string, List<EndesaEntity.Tunel>>();

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

                    f = 1; // Porque la primera fila es la cabecera
                    for (int i = 1; i < 100000000; i++)
                    {
                        c = 1;
                        f++;

                        if (workSheet.Cells[f, 1].Value == null)
                            break;

                        if (workSheet.Cells[f, 1].Value.ToString() == "")
                            break;


                        firstOnly = true;
                        for(int j = 1; j <= 12; j++)
                        {
                            c++;
                            EndesaEntity.Tunel t = new EndesaEntity.Tunel();

                            if (workSheet.Cells[f, 1].Value.ToString().Length > 20)
                                t.cups20 = Convert.ToString(workSheet.Cells[f, 1].Value).Substring(0, 20);
                            else
                                t.cups20 = Convert.ToString(workSheet.Cells[f, 1].Value);

                            if (workSheet.Cells[f, c].Value != null)
                                if (workSheet.Cells[f, c].Value.ToString() != "")
                                {
                                    t.total_energia = Convert.ToDouble(workSheet.Cells[f, c].Value);

                                    anio = Convert.ToInt32(Convert.ToString(workSheet.Cells[1, c].Value).Substring(0, 4));
                                    mes = Convert.ToInt32(Convert.ToString(workSheet.Cells[1, c].Value).Substring(4, 2));

                                    t.startDate = new DateTime(anio, mes, 1);
                                    t.endDate = new DateTime(anio, mes, DateTime.DaysInMonth(anio, mes));


                                    if (!dic_consumos.TryGetValue(t.cups20, out o))
                                    {
                                        o = new List<EndesaEntity.Tunel>();
                                        o.Add(t);
                                        dic_consumos.Add(t.cups20, o);
                                    }
                                    else
                                        o.Add(t);
                                }
                                   
                        }    

                    }
                    
                }
                fs = null;
                excelPackage = null;


                foreach(KeyValuePair<string, EndesaEntity.TunelContrato> p in dic_contratos)
                    foreach(EndesaEntity.Tunel pp in p.Value.lista)
                    {
                        if(dic_consumos.TryGetValue(pp.cups20, out o))
                        {
                            ActualizaConsumos(pp.cups20, pp.startDate, pp.endDate,
                                GetTotalConsumo(o,pp.cups20, pp.startDate, pp.endDate));
                        }
                    }

                ActualizaContratosResumen();


            }
            catch (Exception e)
            {
                fs = null;
                excelPackage = null;

                MessageBox.Show(e.Message,
                    "Estructura de columnas Excel incorrecta",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                
            }

        }

        private double GetTotalConsumo(List<EndesaEntity.Tunel> lista, string cups20, DateTime fd, DateTime fh)
        {
            List<EndesaEntity.Tunel> datos;
            double total_consumo = 0;            
            
            datos = lista.Where(z => z.startDate >= fd && z.endDate <= fh).ToList();
            foreach (EndesaEntity.Tunel p in datos)
                total_consumo += p.total_energia;                

            return total_consumo;
        }

        private void ActualizaConsumos(string cups20, DateTime fd, DateTime fh, double consumo_total)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update tunel_contratos set"
                 + " last_update_by = '" + System.Environment.UserName.ToUpper() + "'";

            strSql += " ,consumo_total = " + consumo_total.ToString().Replace(",", ".");
            

            strSql += " where cups20 = '" + cups20 + "' and"
                + " start_date = '" + fd.ToString("yyyy-MM-dd") + "' and"
                + " end_date = '" + fh.ToString("yyyy-MM-dd") + "'";

            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }
       


        public void InformeExcel()
        {
            int c = 1;
            int f = 1;
            SaveFileDialog save;

            try
            {

                //if(dic_contratos.Count > 0)
                //{
                //    save = new SaveFileDialog();
                //    save.Title = "Ubicación del informe";
                //    save.AddExtension = true;
                //    save.DefaultExt = "xlsx";
                //    DialogResult result = save.ShowDialog();
                //    if (result == DialogResult.OK)
                //    {
                //        FileInfo fileInfo = new FileInfo(save.FileName);
                //        if (fileInfo.Exists)
                //            fileInfo.Delete();


                //        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                //        ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                //        var workSheet = excelPackage.Workbook.Worksheets.Add("CONTRATOS");
                //        var headerCells = workSheet.Cells[1, 1, 1, 46];
                //        var headerFont = headerCells.Style.Font;

                //        workSheet.View.FreezePanes(2, 1);
                //        headerFont.Bold = true;

                //        workSheet.Cells[f, c].Value = "ID"; 
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "CONTRATO"; 
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;                        

                //        workSheet.Cells[f, c].Value = "INICIO CONTRATO"; 
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "FIN CONTRATO"; 
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "VOLUMEN (GW)"; 
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "CONSUMO TOTAL (GW)"; 
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "LOWERBAND"; 
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "HIGHERBAND"; 
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "COMPLETO"; 
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "PUNTOS TOTALES"; 
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "PUNTOS COMPLETOS";
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        foreach (KeyValuePair<string, EndesaEntity.TunelContrato> p in dic_contratos)
                //        {
                //            f++;
                //            c = 1;
                //            workSheet.Cells[f, c].Value = p.Value.id; c++;
                //            workSheet.Cells[f, c].Value = p.Value.name; c++;

                //            if (p.Value.versionStart > DateTime.MinValue)
                //            {
                //                workSheet.Cells[f, c].Value = p.Value.versionStart;
                //                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                //                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //            }
                //            c++;

                //            if (p.Value.versionEnd > DateTime.MinValue)
                //            {
                //                workSheet.Cells[f, c].Value = p.Value.versionEnd;
                //                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                //                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //            }
                //            c++;

                //            workSheet.Cells[f, c].Value = p.Value.volume;
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                //            c++;
                //            workSheet.Cells[f, c].Value = p.Value.lista.Sum(z => z.total_energia);
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                //            c++;
                //            workSheet.Cells[f, c].Value = p.Value.lowerBand;
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                //            c++;
                //            workSheet.Cells[f, c].Value = p.Value.higherBand;
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                //            c++;                           

                //            workSheet.Cells[f, c].Value = 
                //                (p.Value.lista.Where(z => z.completo).Count() == p.Value.lista.Count) ? "Sí" :"No";
                //            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; 
                //            c++;
                //            workSheet.Cells[f, c].Value = p.Value.lista.Count; 
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                //            c++;
                //            workSheet.Cells[f, c].Value = p.Value.lista.Where(z => z.completo).Count(); 
                //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                //            c++;
                //        }

                //        var allCells = workSheet.Cells[1, 1, f, 13];
                //        workSheet.Cells["A1:K1"].AutoFilter = true;


                //        headerFont.Bold = true;
                //        allCells.AutoFitColumns();

                //        workSheet = excelPackage.Workbook.Worksheets.Add("CUPS");
                //        headerCells = workSheet.Cells[1, 1, 1, 46];
                //        headerFont = headerCells.Style.Font;

                //        c = 1;
                //        f = 1;
                //        workSheet.View.FreezePanes(2, 1);
                //        headerFont.Bold = true;
                //        workSheet.Cells[f, c].Value = "ID";
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "CONTRATO";
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "INICIO CONTRATO";
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "FIN CONTRATO";
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "CUPS";
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "FECHA INICIO";
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "FECHA FIN";
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "ÚLTIMA FECHA FACTURA HASTA";
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "CONSUMO TOTAL (kW)";
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;

                //        workSheet.Cells[f, c].Value = "COMPLETO";
                //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //        c++;


                //        foreach (KeyValuePair<string, EndesaEntity.TunelContrato> p in dic_contratos)
                //        {
                //            foreach (EndesaEntity.Tunel q in p.Value.lista)
                //            {
                //                f++;
                //                c = 1;
                //                workSheet.Cells[f, c].Value = p.Value.id; c++;
                //                workSheet.Cells[f, c].Value = p.Value.name; c++;

                //                if (p.Value.versionStart > DateTime.MinValue)
                //                {
                //                    workSheet.Cells[f, c].Value = p.Value.versionStart;
                //                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                //                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //                }
                //                c++;

                //                if (p.Value.versionEnd > DateTime.MinValue)
                //                {
                //                    workSheet.Cells[f, c].Value = p.Value.versionEnd;
                //                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                //                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //                }
                //                c++;

                //                workSheet.Cells[f, c].Value = q.cups20; c++;

                //                if (q.startDate > DateTime.MinValue)
                //                {
                //                    workSheet.Cells[f, c].Value = q.startDate;
                //                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                //                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //                }
                //                c++;

                //                if (q.endDate > DateTime.MinValue)
                //                {
                //                    workSheet.Cells[f, c].Value = q.endDate;
                //                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                //                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //                }
                //                c++;

                //                if (q.ultima_factura > DateTime.MinValue)
                //                {
                //                    workSheet.Cells[f, c].Value = q.ultima_factura;
                //                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                //                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //                }
                //                c++;

                //                workSheet.Cells[f, c].Value = q.total_energia * 1000000;
                //                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                //                c++;

                //                workSheet.Cells[f, c].Value = q.completo ? "Sí" : "No";
                //                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //                c++;


                //            }
                //        }
                            

                //        allCells = workSheet.Cells[1, 1, f, 10];
                //        workSheet.Cells["A1:J1"].AutoFilter = true;


                //        headerFont.Bold = true;
                //        allCells.AutoFitColumns();

                //        excelPackage.Save();

                //        MessageBox.Show("Informe terminado.",
                //               "Informe Excel",
                //               MessageBoxButtons.OK,
                //               MessageBoxIcon.Information);

                //        DialogResult result3 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                //        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                //        if (result3 == DialogResult.Yes)
                //            System.Diagnostics.Process.Start(save.FileName);


                //    }
                //}

            }catch(Exception e)
            {

            }
        }

        public void Save()
        {
            EndesaEntity.TunelContrato inv = new EndesaEntity.TunelContrato();
            inv = dic_contratos.Where(z => z.Key == this.cliente 
            + this.fecha_inicio_tunel.ToString("yyyyMMdd") + this.fecha_final_tunel.ToString("yyyyMMdd")).SingleOrDefault().Value;

            if (inv != null)
                Update();
            else
                New();

        }

        private void New()
        {

        }

        private void Update()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update tunel_contratos_resumen set"
                 + " last_update_by = '" + System.Environment.UserName.ToUpper() + "'";

            if (this.formula_antigua)
                strSql += " ,formula_antigua = " + "'S'";
            else
                strSql += " ,formula_antigua = " + "'N'";

            strSql += " where cliente = '" + this.cliente + "'";    

            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            MessageBox.Show("Se ha modifcado correctamente el contrato: " + this.cliente,
            "Contratos tunel",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
        }

    }
}
