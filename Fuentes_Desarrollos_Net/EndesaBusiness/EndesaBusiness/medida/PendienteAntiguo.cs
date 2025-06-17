using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class PendienteAntiguo
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "PendienteAntiguo");
        public PendienteAntiguo()
        {
            
        }

        public void LanzaProceso()
        {
            //ActualizaPdteWeb_ArchivoActual();
            Archivo_Actual_Total_Ordenado();
        }

        private void ActualizaPdteWeb_ArchivoActual()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                ficheroLog.Add("ActualizaPdteWeb_ArchivoActual");
                ficheroLog.Add("==============================");
                Console.WriteLine("ActualizaPdteWeb_ArchivoActual");
                Console.WriteLine("==============================");
                strSql = "UPDATE segpend_archivo_actual ac"
                    + " INNER JOIN scea s ON"
                    + " s.IDU = ac.ps"
                    + " SET ac.pdte_web = s.PrimerMesPDTE";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch(Exception e)
            {

                ficheroLog.AddError("ActualizaPdteWeb_ArchivoActual: " + e.Message);
            }
        }

        private void Archivo_Actual_Total_Ordenado()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            string importe_texto = "";
            string cups = "";
            string pdte_real = "";
            int numReg = 0;
            int datosTratados = 0;

            try
            {
                ficheroLog.Add("Archivo_Actual_Total_Ordenado");
                ficheroLog.Add("=============================");
                Console.WriteLine("Archivo_Actual_Total_Ordenado");
                Console.WriteLine("=============================");
                strSql = "delete from segpend_archivo_actual_ordenado";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "SELECT" 
                    + " ac.linea_negocio, ac.situacion, ac.pdte_real, ac.pdte_web,"
                    + " s.Subestado, ac.ps, ac.nombre, ac.empresa, ac.area, ac.a_quien," 
                    + " ac.comentario_contratacion, ac.comentario_medida, ac.comentario_facturacion,"
                    + " ac.numero_incidencia, ac.origen, ac.fecha_resolucion,"
                    + " ac.linea_negocio, s.TAM_por_CUPS, ac.cups20, ac.contrato, ac.tipo,"
                    + " ac.propiedad, s.esPrimeraFactura"
                    + " FROM segpend_archivo_actual ac"
                    + " INNER JOIN scea s ON"
                    + " s.IDU = ac.ps"
                    + " ORDER BY ac.linea_negocio, ac.situacion, ac.pdte_real, ac.pdte_web, ac.ps";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    numReg++;
                    datosTratados++;

                    if (firstOnly)
                    {
                        sb.Append("replace into segpend_archivo_actual_ordenado");
                        sb.Append(" (linea_negocio, situacion, pdte_real, pdte_web, ps,");
                        sb.Append("nombre, empresa, area, a_quien, comentario_contratacion,");
                        sb.Append("comentario_medida, comentario_facturacion, numero_incidencia,");
                        sb.Append("origen, fecha_resolucion, importe_pendiente, cups20,");
                        sb.Append("contrato, tipo, es_primera_factura, propiedad) values ");
                        firstOnly = false;
                    }

                    #region campos

                    if (r["linea_negocio"] != System.DBNull.Value)
                        sb.Append("('").Append(r["linea_negocio"].ToString()).Append("',");
                    else
                        sb.Append("(").Append("null").Append(",");

                    if(r["situacion"] != System.DBNull.Value)
                        sb.Append("'").Append(r["situacion"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r["pdte_real"] != System.DBNull.Value)
                    {
                        sb.Append("'").Append(r["pdte_real"].ToString()).Append("',");
                        pdte_real = r["pdte_real"].ToString();
                    }
                        
                    else
                        sb.Append("null").Append(",");

                    if (r["pdte_web"] != System.DBNull.Value)
                        sb.Append("'").Append(r["pdte_web"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r["ps"] != System.DBNull.Value)
                    {
                        sb.Append("'").Append(r["ps"].ToString()).Append("',");
                        cups = r["ps"].ToString();
                    }
                    else
                    {
                        sb.Append("null").Append(",");
                        cups = null;
                    }
                        

                    if (r["nombre"] != System.DBNull.Value)
                        sb.Append("'").Append(r["nombre"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r["empresa"] != System.DBNull.Value)
                        sb.Append("'").Append(r["empresa"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r["area"] != System.DBNull.Value)
                        sb.Append("'").Append(r["area"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r["a_quien"] != System.DBNull.Value)
                        sb.Append("'").Append(r["a_quien"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r["comentario_contratacion"] != System.DBNull.Value)
                        sb.Append("'").Append(r["comentario_contratacion"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r["comentario_medida"] != System.DBNull.Value)
                        sb.Append("'").Append(r["comentario_medida"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r["comentario_facturacion"] != System.DBNull.Value)
                        sb.Append("'").Append(r["comentario_facturacion"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r["numero_incidencia"] != System.DBNull.Value)
                        sb.Append("'").Append(r["numero_incidencia"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r[""] != System.DBNull.Value)
                        sb.Append("'").Append(r[""].ToString()).Append("',");
                    else                        
                        sb.Append("null").Append(",");

                    if (r["fecha_resolucion"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fecha_resolucion"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null").Append(",");

                    

                    importe_texto = "";
                    // Importe pendiente
                    if (r["linea_negocio"].ToString() == "GAS")
                        importe_texto = "GAS";
                    else if(cups != null)
                    {
                        if (cups.Substring(0, 3) == "XXX")
                            importe_texto = "INTERNACIONAL";
                        else if (r["TAM_por_CUPS"] != System.DBNull.Value)
                            importe_texto = Importe_TAM(Convert.ToDouble(r["TAM_por_CUPS"]), 
                                DIF_AAAAMM(DateTime.Now.ToString("yyyyMM"), pdte_real)).ToString();
                        else
                            importe_texto = "NUEVO";
                    }

                    sb.Append("'").Append(importe_texto).Append("',");

                    if (r["cups20"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cups20"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r["contrato"] != System.DBNull.Value)
                        sb.Append("'").Append(r["contrato"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r["tipo"] != System.DBNull.Value)
                        sb.Append("'").Append(r["tipo"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r["es_primera_factura"] != System.DBNull.Value)
                        sb.Append("'").Append(r["es_primera_factura"].ToString()).Append("',");
                    else
                        sb.Append("null").Append(",");

                    if (r["propiedad"] != System.DBNull.Value)
                        sb.Append("'").Append(r["propiedad"].ToString()).Append("'),");
                    else
                        sb.Append("null").Append("),");


                    #endregion

                    if (numReg == 100)
                    {                        
                        Console.WriteLine("Guardando datos. " + (datosTratados + 1) + " datos guardados ...");
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.MED);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        numReg = 0;
                    }

                }
                db.CloseConnection();

                if (numReg > 0)
                {
                    Console.WriteLine("Guardando datos. " + (datosTratados + 1) + " datos guardados ...");
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    numReg = 0;
                }



            }
            catch (Exception e)
            {

                ficheroLog.AddError("Archivo_Actual_Total_Ordenado: " + e.Message);
            }
        }

        private int DIF_AAAAMM(string f1, string f2)
        {
            int dif = 1;

            if (f1 == "" || f2 == "")
                return dif;
            else
                return (Convert.ToInt32(f1) - Convert.ToInt32(f2));

            
        }

        private double Importe_TAM(double tam, int dif)
        {
            return Math.Round(tam * dif,0);
        }


    }
}
