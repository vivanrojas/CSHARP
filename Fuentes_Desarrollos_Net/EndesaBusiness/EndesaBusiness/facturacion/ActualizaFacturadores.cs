using EndesaBusiness.servidores;
using Microsoft.Office.Interop.Word;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;
using System;
using System.Collections.Generic;
using System.Diagnostics.PerformanceData;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{
    public class ActualizaFacturadores
    {


        Dictionary<string, string> dic_cups;

        public ActualizaFacturadores()
        {

        }


        public void LoadDataFromExcelProgramas(string fichero)
        {

            MySQLDB db;
            MySqlCommand command;

            int mercado = 0;
            int c = 1;
            int f = 1;
            string ccounips = "";
            string fecha_temp = "";
            DateTime fecha = new DateTime();
            int anio = 0;
            int mes = 0;
            int dia = 0;
            bool firstOnly = true;
            bool firstOnlyQuery = true;
            StringBuilder sb = new StringBuilder();
            int numReg = 0;
            int totalReg = 0;
            int numPeriodos = 0;


            calendarios.UtilidadesCalendario utilCal = new calendarios.UtilidadesCalendario();

            try
            {

                dic_cups = CargaCupsPrograma();



                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fs);
                int total_hojas_excel = excelPackage.Workbook.Worksheets.Count();
                var workSheet = excelPackage.Workbook.Worksheets.First();


                for (int hoja = 0; hoja < total_hojas_excel; hoja++)
                {
                    workSheet = excelPackage.Workbook.Worksheets[hoja];

                    if (hoja == 0)
                        mercado = 0;
                    else
                        mercado = Convert.ToInt32(workSheet.Name.Trim().Replace("Intra", ""));


                    f = 1; // Porque la primera fila es la cabecera
                    for (int i = 1; i < 100000000; i++)
                    {

                       

                        c = 1;
                        f++;

                        if (workSheet.Cells[f, 1].Value == null)
                            break;

                        if (workSheet.Cells[f, 1].Value.ToString() == "")
                            break;

                        numReg++;
                        totalReg++;


                        ccounips = workSheet.Cells[f, 1].Value.ToString();c++;
                        fecha_temp = workSheet.Cells[f, 2].Value.ToString();
                        anio = Convert.ToInt32(fecha_temp.Substring(0, 4));
                        mes = Convert.ToInt32(fecha_temp.Substring(4, 2));
                        dia = Convert.ToInt32(fecha_temp.Substring(6, 2));
                        fecha = new DateTime(anio, mes, dia);

                        // Actualizamos la tabla Programas Cliente
                        if(!ExisteCUPSPrograma(ccounips))
                            ActualizaProgramaCliente(ccounips);

                        if (firstOnlyQuery)
                        {
                            sb.Append("REPLACE into ag_programas (CCOUNIPS, Fecha, Mercado, Unidad");
                            for (int j = 1; j <= 25; j++)
                                sb.Append(",Value" + j);
                            sb.Append(") values ");
                            firstOnlyQuery = false;
                        }
                        
                        sb.Append("('").Append(ccounips).Append("',");
                        sb.Append("'").Append(fecha.ToString("yyyy-MM-dd")).Append("',");
                        sb.Append(mercado).Append(",");
                        sb.Append("'kWh'");

                        numPeriodos = utilCal.NumPeriodosHorarios(fecha);

                        firstOnly = true;
                        for (int j = 1; j <= numPeriodos; j++)
                        {
                            if (numPeriodos == 25 && j > 3)
                            {
                                if (firstOnly)
                                {
                                    // strSql += " ,Value" + j + " = 0 ";
                                    sb.Append(",").Append(0);
                                    firstOnly = false;
                                }
                                else
                                    //strSql += " ,Value" + j + " = "
                                    //    + workSheet.Cells[f, j + 1].Value.ToString().Replace(",", ".") + " ";
                                    sb.Append(",").Append(workSheet.Cells[f, j + 1].Value.ToString().Replace(",", "."));
                            }
                            else
                            {
                                sb.Append(",").Append(workSheet.Cells[f, j + 2].Value.ToString().Replace(",", "."));
                                
                            }
                                //strSql += " ,Value" + j + " = "
                                //      + workSheet.Cells[f, j + 2].Value.ToString().Replace(",", ".") + " ";
                                
                        }
                        if(numPeriodos == 24)
                            sb.Append(",null");

                        if (numPeriodos == 23)
                            sb.Append(",null,null");

                        sb.Append("),");

                        //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        //command = new MySqlCommand(strSql, db.con);
                        //command.ExecuteNonQuery();
                        //db.CloseConnection();

                        if (numReg == 250)
                        {
                            firstOnlyQuery = true;
                            db = new MySQLDB(MySQLDB.Esquemas.FAC);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            numReg = 0;
                        }


                    }

                    if (numReg > 0)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        numReg = 0;
                    }

                }


                MessageBox.Show("Proceso Finalizado correctamente."
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "Se han importado " + totalReg + " registros."
                        + System.Environment.NewLine,
                  "Importación ficheros Excel Programas Consumo",
                  MessageBoxButtons.OK, MessageBoxIcon.Information);


            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
                  "Importación ficheros Excel Programas Consumo",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        private void ActualizaProgramaCliente(string ccounips)
        {
            MySQLDB db;
            MySqlCommand command;            
            string strSql;
            MySqlDataReader r;
            bool existe = false;

            try
            {

                strSql = "SELECT CCOUNIPS from ag_programascliente where"
                    + " CCOUNIPS = '" + ccounips + "'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                existe = r.Read();
                db.CloseConnection();
                if (!existe)
                {
                    strSql = "REPLACE into ag_programascliente SET "
                        + " CUPSREE = '" + GetCUPSREE(ccounips) + "',"
                        + " CCOUNIPS = '" + ccounips + "'";                        
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
                  "ActualizaFacturadores - ActualizaProgramaCliente",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
            
        }

        private string GetCUPSREE(string ccounips)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string cups20 = "";
                

            strSql = "select CUPS20 from RELACION_CUPS where"
                + " CUPS_CORTO = '" + ccounips + "'";
            db = new MySQLDB(MySQLDB.Esquemas.AUX);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                cups20 = r["CUPS20"].ToString();
            }           
            db.CloseConnection();
            return cups20;
        }

        private Dictionary<string, string> CargaCupsPrograma()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string cups13 = "";

            Dictionary<string, string> d = new Dictionary<string, string>();

            strSql = "SELECT ccounips FROM ag_programascliente GROUP BY ccounips";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {

                if (r["ccounips"] != System.DBNull.Value)
                    cups13 = r["ccounips"].ToString();

                d.Add(cups13, cups13);
            }
            db.CloseConnection();

            return d;


        }

        private bool ExisteCUPSPrograma(string cups)
        {
            string o;
            return dic_cups.TryGetValue(cups, out o);
        }

    }
}
