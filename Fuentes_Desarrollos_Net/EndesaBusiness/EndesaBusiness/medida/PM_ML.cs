using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.servidores;
using System.Data.OleDb;
using System.IO;

namespace EndesaBusiness.medida
{
    public class PM_ML
    {
        EndesaBusiness.utilidades.Param p;
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_PM_ML");
        public PM_ML()
        {
            p = new EndesaBusiness.utilidades.Param("scea_pmml_param", EndesaBusiness.servidores.MySQLDB.Esquemas.MED);
        }


        public void EjecutaProceso()
        {

            bool hayError = false;

            office365.MS_Access msAccess = new office365.MS_Access();
            EndesaEntity.cola.ProcesoCola macro_actualiza_pmml = new EndesaEntity.cola.ProcesoCola();
            EndesaEntity.cola.ProcesoCola consulta_hoja_pmml = new EndesaEntity.cola.ProcesoCola();

            macro_actualiza_pmml.ruta = p.GetValue("ruta_access_macro_actualiza_pmml");
            macro_actualiza_pmml.bbdd = p.GetValue("bbdd_macro_actualiza_pmml");
            macro_actualiza_pmml.nombre_proceso = p.GetValue("macro_actualiza_pmml");
            macro_actualiza_pmml = msAccess.EjecutaMacro(macro_actualiza_pmml);

            if(macro_actualiza_pmml.mensaje_error != null)
            {
                hayError = true;
                GeneraCorreoError(macro_actualiza_pmml.mensaje_error);
            }           

            if (!hayError)
            {
                string archivo = ExportaTablaAccess();



                if (archivo != "" && (p.GetValue("subir_a_FTP") == "S"))
                    Subir_FTP(archivo);

                if ((p.GetValue("mail_enviar") == "S"))
                {
                    GeneraCorreo();
                }
            }
                       

        }

        public void UpdateFechaCopia()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update scea_pmml_param"
                + " set value = '" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "'"
                + " where code = 'ultima_exportacion'";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
        }



        private string ExportaTablaAccess()
        {
            string archivo = "";
            string strSql = "";
            EndesaBusiness.servidores.AccessDB ac;
            OleDbDataReader r;
            StringBuilder sb = new StringBuilder();
            string linea = "";
            bool firstOnly = true;

            try
            {

                string[] listaArchivos = Directory.GetFiles(p.GetValue("salida_programada"), p.GetValue("prefijo_archivo") + "*.csv");
                for (int i = 0; i < listaArchivos.Length; i++)
                {
                    ficheroLog.Add("Borramos el archivo previo: " + listaArchivos[i]);
                    FileInfo file = new FileInfo(listaArchivos[i]);
                    file.Delete();
                }

                archivo = p.GetValue("salida_programada") +
                    p.GetValue("prefijo_archivo") +
                    "_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";

                StreamWriter swa = new StreamWriter(archivo, false);
                             


                strSql = "SELECT[PM ML].CUPS AS CUPS13, [PM ML].PF, [PM ML].PM AS CUPS15PM," +
                    " IIf([CUPS22]= 'ES','',[CUPS22]) AS CUPS22PM, SCEA.[Descripcion Estado SCE], " +
                    "SCEA.tdistri, [PM ML].Contrato, [PM ML].nPMs, [PM ML].[nPMs P], [PM ML].TIPOPM, " +
                    "[PM ML].VERSION, [PM ML].[TLF ML], [PM ML].[DE ML], [PM ML].[PM ML], Val([CL ML]) AS CL_ML, " +
                    "[PM ML].[TLF AUX], [PM ML].[IP ML], [PM ML].[PUERTO ML], [PM ML].MODO_LECT" +
                    " FROM[PM ML] LEFT JOIN SCEA ON[PM ML].CUPS = SCEA.IDU";

                strSql = strSql.Replace("'", "\"");

                ficheroLog.Add("Ejecutando la consulta: " + strSql);

                ac = new EndesaBusiness.servidores.AccessDB(p.GetValue("access_ruta"));
                OleDbCommand cmd = new OleDbCommand(strSql, ac.con);
                r = cmd.ExecuteReader();
                while (r.Read())
                {
                    if (firstOnly)
                    {
                        ficheroLog.Add("Exportando datos al archivo: " + archivo);
                        linea = "CUPS13;PF;CUPS15PM;CUPS22PM;Descripcion Estado SCE;tdistri;Contrato;nPMs;nPMs P;" +
                            "TIPOPM;VERSION;TLF ML;DE ML;PM ML;CL_ML;TLF AUX;IP ML;PUERTO ML;MODO_LECT";
                        firstOnly = false;
                    }
                    else
                    {
                        if (r["CUPS13"] != null)
                            linea = r["CUPS13"].ToString();
                        linea = linea + ";";

                        if (r["PF"] != null)
                            linea = linea + r["PF"].ToString();
                        linea = linea + ";";

                        if (r["CUPS15PM"] != null)
                            linea = linea + r["CUPS15PM"].ToString();
                        linea = linea + ";";

                        if (r["CUPS22PM"] != null)
                            linea = linea + r["CUPS22PM"].ToString();
                        linea = linea + ";";

                        if (r["Descripcion Estado SCE"] != null)
                            linea = linea + r["Descripcion Estado SCE"].ToString();
                        linea = linea + ";";

                        if (r["tdistri"] != null)
                            linea = linea + r["tdistri"].ToString();
                        linea = linea + ";";

                        if (r["Contrato"] != null)
                            linea = linea + r["Contrato"].ToString();
                        linea = linea + ";";

                        if (r["nPMs"] != null)
                            linea = linea + r["nPMs"].ToString();
                        linea = linea + ";";

                        if (r["nPMs P"] != null)
                            linea = linea + r["nPMs P"].ToString();
                        linea = linea + ";";

                        if (r["TIPOPM"] != null)
                            linea = linea + r["TIPOPM"].ToString();
                        linea = linea + ";";

                        if (r["VERSION"] != null)
                            linea = linea + r["VERSION"].ToString();
                        linea = linea + ";";

                        if (r["TLF ML"] != null)
                            linea = linea + r["TLF ML"].ToString();
                        linea = linea + ";";

                        if (r["DE ML"] != null)
                            linea = linea + r["DE ML"].ToString();
                        linea = linea + ";";

                        if (r["PM ML"] != null)
                            linea = linea + r["PM ML"].ToString();
                        linea = linea + ";";

                        if (r["CL_ML"] != null)
                            linea = linea + r["CL_ML"].ToString();
                        linea = linea + ";";

                        if (r["TLF AUX"] != null)
                            linea = linea + r["TLF AUX"].ToString();
                        linea = linea + ";";

                        if (r["IP ML"] != null)
                            linea = linea + r["IP ML"].ToString();
                        linea = linea + ";";

                        if (r["PUERTO ML"] != null)
                            linea = linea + r["PUERTO ML"].ToString();
                        linea = linea + ";";

                        if (r["MODO_LECT"] != null)
                            linea = linea + r["MODO_LECT"].ToString();
                    }



                    swa.WriteLine(linea);
                }
                ac.CloseConnection();
                swa.Close();

                return archivo;
            }
            catch (Exception e)
            {

                ficheroLog.AddError(" Exportación tabla PM ML a FTP " + e.Message);
                return "";

            }


        }

        private void Subir_FTP(string archivo)
        {

            EndesaBusiness.utilidades.UltimateFTP ftp;
            FileInfo file;
            EndesaBusiness.medida.PM_ML pmml;

            try
            {

                pmml = new EndesaBusiness.medida.PM_ML();
                file = new FileInfo(archivo);

                ftp = new EndesaBusiness.utilidades.UltimateFTP(
                        p.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                        p.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                        EndesaBusiness.utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_pass", DateTime.Now, DateTime.Now), true),
                        p.GetValue("ftp_port", DateTime.Now, DateTime.Now));

                ftp.Upload(p.GetValue("ftp_ruta_salida") + file.Name, archivo);
                this.UpdateFechaCopia();

                p = new EndesaBusiness.utilidades.Param("scea_pmml_param", EndesaBusiness.servidores.MySQLDB.Esquemas.MED);

                ficheroLog.Add("Se ha exportado el archivo "
                   + archivo + " correctamente.");

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Exportación tabla PM ML a FTP: " + e.Message);                
            }
        }


        private void GeneraCorreo()
        {
            string body = "";
            string subject = "";
            string from = "";
            string to = "";
            string cc = null;
            string attachment = null;


            try
            {
                from = p.GetValue("mail_from");
                to = p.GetValue("mail_to");
                subject = p.GetValue("mail_subject");
                cc = p.GetValue("mail_cc");

               
                body = (DateTime.Now.Hour > 14 ? "Buenas tardes." : "Buenos días.")
                    + System.Environment.NewLine
                    + "     Se han generado el archivo correctamente."
                    + System.Environment.NewLine
                    + "Un saludo.";


                // EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
                mes.SendMail(from, to, cc, subject, body, attachment);

            }
            catch (Exception e)
            {
                ficheroLog.AddError("GeneraCorreo: " + e.Message);
            }
        }

        private void GeneraCorreoError(string error)
        {
            string body = "";
            string subject = "";
            string from = "";
            string to = "";
            string cc = null;
            string attachment = null;


            try
            {
                from = p.GetValue("mail_from");
                to = p.GetValue("mail_to");
                subject = "Error en proceso Extracción PM ML Kronos";
                cc = p.GetValue("mail_cc");


                body = (DateTime.Now.Hour > 14 ? "Buenas tardes." : "Buenos días.")
                    + System.Environment.NewLine
                    + "     Se ha generado el siguiente error en el proceso:"
                    + System.Environment.NewLine
                    + error
                    + System.Environment.NewLine
                    + "Un saludo.";


                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
                mes.SendMail(from, to, cc, subject, body, attachment);

            }
            catch (Exception e)
            {
                ficheroLog.AddError("GeneraCorreo: " + e.Message);
            }
        }


    }


}
