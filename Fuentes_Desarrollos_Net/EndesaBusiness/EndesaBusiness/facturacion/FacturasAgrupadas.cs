using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{
    public class FacturasAgrupadas
    {
        public FacturasAgrupadas()
        {

        }

        public void CopiaFacturasAgrupadasOWEN()
        {

            OracleServer ora_db;
            OracleCommand ora_command;
            OracleDataReader r;
            MySqlDataReader rr;

            MySQLDB db;
            MySqlCommand command;

            string strSql = "";
            StringBuilder sb = new StringBuilder();
            int i = 0;
            int registros_totales = 0;
            bool firstOnly = true;
            long n;
            DateTime ultimaFecha = new DateTime();

            try
            {

                //strSql = "SELECT MAX(FFACTDES) as FFACTDES FROM fo_agrupadas";
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                //rr = command.ExecuteReader();
                //while (rr.Read())
                //{                    
                //    ultimaFecha = Convert.ToDateTime(rr["FFACTDES"]);
                //    Console.WriteLine("Ultima fecha guardada en fo_agrupadas = "
                //        + ultimaFecha.ToString("dd/MM/yyyy"));
                //}
                //db.CloseConnection();

                //ultimaFecha = new DateTime(2021, 01, 01);
                ultimaFecha = new DateTime();
                ultimaFecha = DateTime.Now;

                strSql = "SELECT CREFEREN, SECFACTU, TESTFACT,"
                     + " FFACTDES, FFACTHAS,"
                     + " CREFAGP, CSECAGP, CFACAGP"
                     + " FROM OWEN_OWNER.FACT_DETALLEOPERACIONES WHERE"
                     + " TFACTURA = 3 AND"
                     + " FFACTDES >= '" + ultimaFecha.AddDays(-10).ToString("dd/MM/yyyy") + "' AND"
                     + " CREFEREN is NOT NULL AND"
                     + " CFACAGP IS NOT NULL";
                Console.WriteLine(strSql);
                ora_db = new OracleServer(OracleServer.Servidores.OWE);
                ora_command = new OracleCommand(strSql, ora_db.con);
                Console.WriteLine("Consultando BBDD ...");
                r = ora_command.ExecuteReader();
                Console.WriteLine("Copiando datos ...");
                while (r.Read())
                {
                 
                    if (!r["CREFEREN"].ToString().Contains("E") && 
                         long.TryParse(r["CREFAGP"].ToString(), out n) &&
                         long.TryParse(r["CSECAGP"].ToString(), out n))
                    {

                        i++;
                        registros_totales++;

                        if (firstOnly)
                        {
                            sb.Append("replace into fo_agrupadas_aux");
                            sb.Append(" (CREFEREN, SECFACTU, TESTFACT,");
                            sb.Append(" FFACTDES, FFACTHAS,");
                            sb.Append(" CREFAGP, CSECAGP, CFACAGP) values ");
                            firstOnly = false;
                        }

                        sb.Append("(").Append(Convert.ToInt64(r["CREFEREN"])).Append(",");
                        sb.Append(Convert.ToInt32(r["SECFACTU"])).Append(",");
                        sb.Append("'").Append(r["TESTFACT"].ToString()).Append("',");

                        if (r["FFACTDES"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["FFACTDES"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["FFACTHAS"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["FFACTHAS"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["CREFAGP"] != System.DBNull.Value)
                            sb.Append(Convert.ToInt64(r["CREFAGP"])).Append(",");
                        else
                            sb.Append("null,");

                        if (r["CSECAGP"] != System.DBNull.Value)
                            sb.Append(Convert.ToInt64(r["CSECAGP"])).Append(",");
                        else
                            sb.Append("null,");

                        sb.Append("'").Append(r["CFACAGP"].ToString()).Append("'),");

                        if (i == 250)
                        {
                            Console.CursorLeft = 0;
                            Console.Write(String.Format("{0:#,##0}", registros_totales));

                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.FAC);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            i = 0;
                        }
                    }
                     
                }
                

                if (i > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    i = 0;
                }

                Console.WriteLine("Fin copia");
                
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
            
        }

        public void CargaExcel(string fichero)
        {

                        
            int f = 1;
            bool firstOnly = true;
            bool firstOnly_agrupada = true;
            int num_fila = 0;
            string fecha = "";

            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            long crefagp = 0;
            int csecagp = 0;
            string cfacagp = "";

            try
            {



                System.IO.FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();


                f = 4; // Leemos a partir de la fila 5
                for (int i = 0; i < 5000; i++)
                {
                    // FORMATO HASTA 20210610
                    // 182 - Nº FISCAL DE LA FACTURA AGRUPADA
                    // 183 - CONTRATO COMERCIAL
                    // 184 - SECUENCIAL FACTURA AGRUPADA

                    // FORMATO DESDE 20210610
                    // 224 - Nº FISCAL DE LA FACTURA AGRUPADA
                    // 225 - CONTRATO COMERCIAL
                    // 226 - SECUENCIAL FACTURA AGRUPADA


                    // FORMATO DESDE 20211109
                    // 225 - Nº FISCAL DE LA FACTURA AGRUPADA
                    // 226 - CONTRATO COMERCIAL
                    // 227 - SECUENCIAL FACTURA AGRUPADA
                    // 162 - CIM


                    f++;
                    if (firstOnly_agrupada && workSheet.Cells[f, 225].Value != null)
                    {
                        crefagp = Convert.ToInt64(workSheet.Cells[f, 226].Value);
                        csecagp = Convert.ToInt32(workSheet.Cells[f, 227].Value);
                        cfacagp = workSheet.Cells[f, 225].Value.ToString().Trim();
                        firstOnly_agrupada = false;
                    }
                }



                f = 4; // Leemos a partir de la fila 5
                for (int i = 0; i < 5000; i++)
                {
                   
                    f++;
                    if (workSheet.Cells[f, 9].Value != null)
                    {
                        num_fila++;
                        if (firstOnly)
                        {
                            sb.Append("replace into fo_agrupadas");
                            sb.Append(" (empresa_id, CREFEREN, SECFACTU, TESTFACT,");
                            sb.Append(" FFACTDES, FFACTHAS,");
                            sb.Append(" CREFAGP, CSECAGP, CFACAGP, CONTRATO_COMERCIAL) values ");
                            firstOnly = false;
                        }                                             


                        sb.Append("(3,").Append(Convert.ToInt64(workSheet.Cells[f, 3].Value)).Append(",");

                        //sb.Append(Convert.ToInt32(workSheet.Cells[f, 132].Value)).Append(",");
                        sb.Append(Convert.ToInt32(workSheet.Cells[f, 175].Value)).Append(",");

                        sb.Append("'").Append((workSheet.Cells[f, 8].Value).ToString()).Append("',");

                        if (workSheet.Cells[f, 9].Value.ToString() != "")
                        {

                            fecha = workSheet.Cells[f, 9].Value.ToString().Substring(0, 4)
                                + "-" + workSheet.Cells[f, 9].Value.ToString().Substring(4, 2)
                                + "-" + workSheet.Cells[f, 9].Value.ToString().Substring(6, 2);
                            sb.Append("'").Append(fecha).Append("',");
                        }
                        else
                            sb.Append("null,");

                        if (workSheet.Cells[f, 10].Value.ToString() != "")
                        {

                            fecha = workSheet.Cells[f, 10].Value.ToString().Substring(0, 4)
                                + "-" + workSheet.Cells[f, 10].Value.ToString().Substring(4, 2)
                                + "-" + workSheet.Cells[f, 10].Value.ToString().Substring(6, 2);
                            sb.Append("'").Append(fecha).Append("',");
                        }
                        else
                            sb.Append("null,");
                        
                        sb.Append(crefagp).Append(",");
                        sb.Append(csecagp).Append(",");
                        sb.Append("'").Append(cfacagp).Append("',");
                        if (workSheet.Cells[f, 226].Value != null)
                            sb.Append(workSheet.Cells[f, 226].Value.ToString()).Append("),");
                        else
                            sb.Append("null),");


                        if (num_fila == 250)
                        {
                            //Console.CursorLeft = 0;
                            //Console.Write(String.Format("{0:#,##0}", registros_totales));

                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.FAC);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            num_fila = 0;
                        }


                    }
                }

                if (num_fila > 0)
                {
                    //Console.CursorLeft = 0;
                    //Console.Write(String.Format("{0:#,##0}", registros_totales));

                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    num_fila = 0;
                }


                fs = null;
                excelPackage = null;

            }
            catch (Exception e)
            {
                MessageBox.Show("Error en el fichero:"
                    + System.Environment.NewLine
                    + fichero
                    + System.Environment.NewLine
                    + e.Message,
                        "Error en el formato del fichero ",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
            }
        }




    }
}
