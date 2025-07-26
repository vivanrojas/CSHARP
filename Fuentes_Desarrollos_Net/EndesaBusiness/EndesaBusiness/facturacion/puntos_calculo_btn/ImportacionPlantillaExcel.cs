using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion.puntos_calculo_btn
{
    public class ImportacionPlantillaExcel
    {

        logs.Log ficheroLog;

        public Dictionary<string, 
             EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> dic_puntos_calculo
        { get; set; }

        public ImportacionPlantillaExcel()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Calculo_Prefacturas_BTN");
            dic_puntos_calculo = new Dictionary<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo>();
        }


        public void CargaExcel(string fichero)
        {
            dic_puntos_calculo = ProcesaExcel(fichero);
            GuardaDatosExcel(dic_puntos_calculo);
        }

        private Dictionary<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> ProcesaExcel(string fichero)
        {
            Dictionary<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> d
                = new Dictionary<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo>();

            int c = 1;
            int f = 1;
            int total_hojas_excel = 0;
            
            
            FileStream fs;
            ExcelPackage excelPackage;
            List<EndesaEntity.Tunel> lista = new List<EndesaEntity.Tunel>();            

            try
            {
                fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                excelPackage = new ExcelPackage(fs);
                total_hojas_excel = excelPackage.Workbook.Worksheets.Count();
                var workSheet = excelPackage.Workbook.Worksheets["Listado Puntos"];

                
               
                

                //if (cabecera != p.GetValue("cabecera_excel", DateTime.Now, DateTime.Now))
                //{
                //    MessageBox.Show("Estructura de columnas en hoja Excel incorrecta."
                //    + System.Environment.NewLine
                //    + System.Environment.NewLine
                //    + "Las columnas del archivo excel no son las esperadas o están en lugares distintos a los esperados.",
                //   "Estructura de columnas Excel incorrecta",
                //   MessageBoxButtons.OK,
                //   MessageBoxIcon.Error);
                //    break;
                //}

                

                f = 1; // Porque la primera fila es la cabecera
                for (int i = 1; i < 100000000; i++)
                {
                    c = 1;
                    f++;

                    if (workSheet.Cells[f, 1].Value == null)
                        break;

                    if (workSheet.Cells[f, 1].Value.ToString() == "")
                        break;

                    EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo t 
                        = new EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo();

                    
                    t.cpe = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    t.contrato = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    t.pct_aplicacion = Convert.ToDouble(workSheet.Cells[f, c].Value) * 100; c++;

                    EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo o;
                    if (!d.TryGetValue(t.cpe, out o))
                        d.Add(t.cpe, t);
                        

                    
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

        private void GuardaDatosExcel(Dictionary<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> dic)
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

                foreach (KeyValuePair<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> p in dic)
                {
                    
                    total_registros_progress_bar++;
                }

                //pb.Text = "Analizando contratos";
                //pb.Show();
                //pb.progressBar.Step = 1;
                //pb.progressBar.Maximum = total_registros_progress_bar + 2;

                progreso++;

                foreach (KeyValuePair<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> p in dic)
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
                        sb.Append("replace into lpc_btn_puntos");
                        sb.Append(" (cpe, contrato, pct_aplicacion, created_by,");
                        sb.Append(" created_date)");
                        sb.Append(" values ");


                        firstOnly = false;
                    }                  
                    
                    sb.Append("('").Append(p.Value.cpe).Append("',");
                    sb.Append(p.Value.contrato).Append(",");                                        
                    sb.Append(p.Value.pct_aplicacion.ToString().Replace(",", ".")).Append(",");
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
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "Tunel.RellenaInventario",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        public void CargaDatos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {

                strSql = "SELECT cpe, contrato, pct_aplicacion"
                  + " from lpc_btn_puntos";                  

                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo c =
                        new EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo();

                    c.cpe = r["cpe"].ToString();
                    c.contrato = r["contrato"].ToString();
                    c.pct_aplicacion = Convert.ToDouble(r["pct_aplicacion"]);                        

                    EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo o;
                    if (!dic_puntos_calculo.TryGetValue(c.cpe, out o))
                        dic_puntos_calculo.Add(c.cpe, c);
                }
                db.CloseConnection();

            }
            catch(Exception ex)
            {
                ficheroLog.AddError("ImportacionPlantillaExcel.CargaDatos: " + ex.Message);
            }
        }

    }
}
