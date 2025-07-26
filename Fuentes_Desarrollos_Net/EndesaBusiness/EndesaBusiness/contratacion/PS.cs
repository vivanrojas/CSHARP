using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion
{
    public class PS
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_PS");
        contratacion.FechasProcesosPS fechasPS = new FechasProcesosPS();
        utilidades.Seguimiento_Procesos ss_pp;


        public PS()
        {
            ss_pp = new utilidades.Seguimiento_Procesos();
        }

        public void DescargaPS()
        {
            string extractor = "";
            DateTime ultimoDiaH = new DateTime();
            utilidades.Global g = new utilidades.Global();            

            EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
            FechasProcesosPS fechasProceso = new FechasProcesosPS();
            bool cargaOK = false;
            StringBuilder textBody = new StringBuilder();
            try
            {                                
                ultimoDiaH = utilidades.Fichero.UltimoDiaHabil();                
                
                ficheroLog.Add(fechasPS.FechaProceso("PS").ToString("dd/MM/yyyy HH:mm")
                + " < "
                + ultimoDiaH.Date.AddHours(23).AddMinutes(59).ToString("dd/MM/yyyy HH:mm")
                + " --> " +
                (fechasPS.FechaProceso("PS") < ultimoDiaH.Date.AddHours(23).AddMinutes(59) ? "true" : "false"));

                Console.WriteLine(fechasPS.FechaProceso("PS").ToString("dd/MM/yyyy HH:mm")
                    + " < "
                    + ultimoDiaH.Date.AddHours(23).AddMinutes(59).ToString("dd/MM/yyyy HH:mm")
                    + " --> " +
                    (fechasPS.FechaProceso("PS") < ultimoDiaH.Date.AddHours(23).AddMinutes(59) ? "true" : "false"));

                                   
                

                if (fechasPS.FechaProceso("PS") < ultimoDiaH.Date.AddHours(23).AddMinutes(59))                                                
                {

                    ss_pp.Update_Fecha_Inicio("Contratación", "PS_AT", "Ejecución Extractor PS");
                    extractor = GetParameter("PS_Extractor");
                    ficheroLog.Add("Ejecutando extractor: " + extractor + " " + ultimoDiaH.ToString("MMdd"));
                    Console.WriteLine("Ejecutando extractor: " + extractor + " " + ultimoDiaH.ToString("MMdd"));
                    utilidades.Fichero.EjecutaComando(extractor, ultimoDiaH.ToString("MMdd"));
                    ficheroLog.Add("Extractor finalizado.");


                    FileInfo archOP =  new FileInfo(GetParameter("PS_ruta")
                        + GetParameter("PS_Prefijo_1") + ultimoDiaH.ToString("MMdd") + ".txt");

                    FileInfo archGP = new FileInfo(GetParameter("PS_ruta")
                        + GetParameter("PS_Prefijo_2") + ultimoDiaH.ToString("MMdd") + ".txt");

                    FileInfo archBTE = new FileInfo(GetParameter("PS_ruta")
                        + GetParameter("NombreFicheroOrigenContrata") + ultimoDiaH.ToString("MMdd") + ".txt");

                    ss_pp.Update_Fecha_Fin("Contratación", "PS_AT", "Ejecución Extractor PS");

                    if (archOP.Exists && archOP.Length > 10440648)                     
                    {
                        

                        cargaOK = CargaTablaPS_OP(archOP.FullName);
                        if (cargaOK)
                        {
                            fechasProceso.ActualizaFechaProceso("PS");                            

                            textBody.Append(System.Environment.NewLine);
                            textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                            textBody.Append(System.Environment.NewLine);
                            textBody.Append("  Se ha realizado el proceso de actualización de la tabla PS.");
                            textBody.Append(System.Environment.NewLine);
                            textBody.Append(System.Environment.NewLine);
                            textBody.Append("Un saludo.");


                            mes.SendMail("contratacionee@enel.com",
                                ListaDestinatarios("Para"),
                                ListaDestinatarios("Copia"),
                                "Tabla PS actualizada " + DateTime.Now.ToString("dd/MM/yyyy"),                              
                                textBody.ToString(), null);

                            EnviaInventarioBTE(archBTE);

                            
                        }
                        else
                        {

                            textBody.Clear();
                            textBody.Append(System.Environment.NewLine);
                            textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                            textBody.Append(System.Environment.NewLine);
                            textBody.Append("  Error en Carga PS");
                            textBody.Append(System.Environment.NewLine);
                            textBody.Append(System.Environment.NewLine);
                            textBody.Append("Un saludo.");

                            mes.SendMail("rsiope.gma@enel.com",
                                "rsiope.gma@enel.com",
                                null,
                                "Error en carga Tabla PS " + DateTime.Now.ToString("dd/MM/yyyy"),                                
                                textBody.ToString(), null);
                        }


                    }
                    else
                    {
                        ss_pp.Update_Comentario("Contratación", "PS_AT", 
                            "Ejecución Extractor PS","No existe el archivo "
                            + archOP.Name);
                    }

                    if(archGP.Exists && archGP.Length > 1000641520)
                    {
                        cargaOK = CargaTablaPS_GP(archGP.FullName);
                    }

                    
                }
                
            }
            catch (Exception e)
            {
                ficheroLog.AddError("DescargaPS: " + e.Message);
            }

        }

        public void DescargaPS_DP()
        {
            DateTime inicio = new DateTime();
            DateTime fin = new DateTime();
            utilidades.Global g = new utilidades.Global();
            try
            {
                ss_pp.Update_Fecha_Inicio("Contratación", "PS D&P", "1_Ejecución Extractor D&P");

                inicio = DateTime.Now;
                utilidades.Fichero.EjecutaComando(GetParameter("PS_Extractor"), utilidades.Fichero.UltimoViernes_MMDD());
                fin = DateTime.Now;
                g.SaveProcess("DescargaPS_DP", "Descarga fichero ps_gp para D&P", inicio, inicio, fin);
                ss_pp.Update_Fecha_Fin("Contratación", "PS D&P", "1_Ejecución Extractor D&P");
            }
            catch (Exception e)
            {
                fin = DateTime.Now;
                g.SaveProcess("DescargaPS_DP", "ERROR: " + utilidades.FuncionesTexto.ArreglaAcentos(e.Message), inicio, inicio, fin);
                ficheroLog.AddError("DescargaPS_DP: " + e.Message);
            }

        }

        public void EnviaPS_DP()
        {
            String prefijoArchivo;
            String ubicacion_origen;
            String ubicacion_destino;
            String ultimoViernes;


            utilidades.ZipUnZip zip = new utilidades.ZipUnZip();
            try
            {
                ultimoViernes = utilidades.Fichero.UltimoViernes_MMDD();

                prefijoArchivo = GetParameter("PS_Prefijo_1");
                ubicacion_origen = GetParameter("DP_Inbox_2");
                ubicacion_destino = GetParameter("DP_Inbox");

                ficheroLog.Add("Comprimiendo archivo " + ubicacion_origen + prefijoArchivo + ultimoViernes + ".txt");
                zip.ComprimirArchivo(ubicacion_origen + prefijoArchivo + ultimoViernes + ".txt",
                    ubicacion_origen + prefijoArchivo + ultimoViernes + ".zip");

                File.Move(ubicacion_origen + prefijoArchivo + ultimoViernes + ".zip",
                        ubicacion_destino + prefijoArchivo + ultimoViernes + ".zip");

                File.Delete(ubicacion_origen + prefijoArchivo + ultimoViernes + ".txt");

                prefijoArchivo = GetParameter("PS_Prefijo_2");
                ficheroLog.Add("Comprimiendo archivo " + ubicacion_origen + prefijoArchivo + ultimoViernes + ".txt");
                zip.ComprimirArchivo(ubicacion_origen + prefijoArchivo + ultimoViernes + ".txt",
                    ubicacion_origen + prefijoArchivo + ultimoViernes + ".zip");

                File.Move(ubicacion_origen + prefijoArchivo + ultimoViernes + ".zip",
                        ubicacion_destino + prefijoArchivo + ultimoViernes + ".zip");


                //                cargaTablaPS_GP(ubicacion_origen + prefijoArchivo + ultimoViernes + ".txt");
                File.Delete(ubicacion_origen + prefijoArchivo + ultimoViernes + ".txt");


            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError(e.Message);
            }


        }
        public void EnviaPS_CEFACO()
        {
            String prefijoArchivo;
            String ubicacion_origen;
            String ubicacion_destino;
            String ultimoViernes;


            utilidades.ZipUnZip zip = new utilidades.ZipUnZip();
            try
            {
                ss_pp.Update_Fecha_Inicio("Contratación", "PS D&P", "2_Envío archivos PS");
                ultimoViernes = utilidades.Fichero.UltimoViernes_MMDD();

                prefijoArchivo = GetParameter("PS_Prefijo_1");
                ubicacion_origen = GetParameter("CEFACO_Inbox_2");
                ubicacion_destino = GetParameter("ruta_sharepoint");

                ficheroLog.Add("Comprimiendo archivo " + ubicacion_origen + prefijoArchivo + ultimoViernes + ".txt");
                zip.ComprimirArchivo(ubicacion_origen + prefijoArchivo + ultimoViernes + ".txt",
                    ubicacion_origen + prefijoArchivo + ultimoViernes + ".zip");

                File.Move(ubicacion_origen + prefijoArchivo + ultimoViernes + ".zip",
                        ubicacion_destino + prefijoArchivo + ultimoViernes + ".zip");

                File.Delete(ubicacion_origen + prefijoArchivo + ultimoViernes + ".txt");

                prefijoArchivo = GetParameter("PS_Prefijo_2");
                ficheroLog.Add("Comprimiendo archivo " + ubicacion_origen + prefijoArchivo + ultimoViernes + ".txt");
                zip.ComprimirArchivo(ubicacion_origen + prefijoArchivo + ultimoViernes + ".txt",
                    ubicacion_origen + prefijoArchivo + ultimoViernes + ".zip");

                File.Move(ubicacion_origen + prefijoArchivo + ultimoViernes + ".zip",
                        ubicacion_destino + prefijoArchivo + ultimoViernes + ".zip");

                
                File.Delete(ubicacion_origen + prefijoArchivo + ultimoViernes + ".txt");

                ss_pp.Update_Fecha_Fin("Contratación", "PS D&P", "2_Envío archivos PS");

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError(e.Message);
            }


        }

        public void CargaFicheroPS(string fileName, int maxLineas)
        {
            StringBuilder sbOP = new StringBuilder();
            StringBuilder sbGP = new StringBuilder();
            Boolean firstOnlyOP = true;
            Boolean firstOnlyGP = true;
            MySQLDB db;
            MySqlCommand command;
            string line = "";
            long i = 0;
            long numLineaGP = 0;
            long numLineaOP = 0;
            long totalLineas = 0;
            string[] campos;

            try
            {

                System.IO.StreamReader file = new System.IO.StreamReader(fileName, System.Text.Encoding.GetEncoding(1252));
                while ((line = file.ReadLine()) != null)
                {

                    totalLineas++;
                    campos = line.Split('|');

                    if (campos[26] == "N")
                    {
                        i = 0;
                        numLineaOP++;

                        if (firstOnlyOP)
                        {
                            sbOP.Append("REPLACE INTO ps_vigentes_gatr_op (EMPRESA, IDU, DDISTRIB, CCONTATR, CNUMCATR, FPREALTA,");
                            sbOP.Append("FALTACON, FPSERCON, FPREVBAJ, TESTCONT, CTARIFA, VTENSIOM,");
                            sbOP.Append("TDISCHOR, CONTREXT, CNUMSCPS, VPOTCAL1, VPOTCAL2, VPOTCAL3,");
                            sbOP.Append("VPOTCAL4, VPOTCAL5, VPOTCAL6, CONSESTI, TINDGCPY, CCONTCOM,");
                            sbOP.Append("TTICONPS, FBAJACON, CSEGMERC, DMUNICIP, DPROVINC, CNIFDNIC,");
                            sbOP.Append("DAPERSOC, CUPSREE, TPERFCNS, TFACTURA, TPUNTMED, TINDTUR,");
                            sbOP.Append("TEQMEDIN, PROVMUNI, CLASESUM, TDIHORAA, CODFINCA, TELEMEDIDA) values ");
                            firstOnlyOP = false;
                        }

                        sbOP.Append("('").Append(campos[i].Trim()).Append("', "); i++; // EMPRESA
                        sbOP.Append("'").Append(campos[i].Trim()).Append("', "); i++; // IDU
                        sbOP.Append("'").Append(campos[i].Trim().Replace("'", "´")).Append("', "); i++; // DDISTRIB
                        sbOP.Append(campos[i].Trim()).Append(", "); i++; // CCONTATR
                        sbOP.Append(campos[i].Trim()).Append(", "); i++; // CNUMCATR

                        for (int x = 0; x <= 9; x++)
                        {
                            sbOP.Append("'").Append(campos[i].Trim()).Append("', "); i++;
                        }

                        for (int x = 0; x <= 6; x++)
                        {
                            sbOP.Append(campos[i].Trim()).Append(", "); i++;
                        }

                        for (int x = 0; x <= 11; x++)
                        {
                            sbOP.Append("'").Append(campos[i].Trim().Replace("'", "´")).Append("', "); i++;
                        }

                        sbOP.Append(campos[i].Trim()).Append(", "); i++; // TPUNTMED

                        for (int x = 0; x <= 5; x++)
                        {
                            sbOP.Append("'").Append(campos[i].Trim()).Append("', "); i++;
                        }

                        sbOP.Append("'").Append(campos[i].Trim()).Append("'),");

                        if (numLineaOP == maxLineas)
                        {
                            Console.WriteLine("Leyendo " + totalLineas + " lineas...");
                            firstOnlyOP = true;
                            db = new MySQLDB(MySQLDB.Esquemas.CON);
                            command = new MySqlCommand(sbOP.ToString().Substring(0, sbOP.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sbOP = null;
                            sbOP = new StringBuilder();
                            numLineaOP = 0;
                        }

                    } //  if (campos[26] == "N")
                    else
                    {
                        i = 0;
                        numLineaGP++;

                        if (firstOnlyGP)
                        {
                            sbGP.Append("REPLACE INTO ps_vigentes_gatr_gp (EMPRESA, IDU, DDISTRIB, CCONTATR, CNUMCATR, FPREALTA,");
                            sbGP.Append("FALTACON, FPSERCON, FPREVBAJ, TESTCONT, CTARIFA, VTENSIOM,");
                            sbGP.Append("TDISCHOR, CONTREXT, CNUMSCPS, VPOTCAL1, VPOTCAL2, VPOTCAL3,");
                            sbGP.Append("VPOTCAL4, VPOTCAL5, VPOTCAL6, CONSESTI, TINDGCPY, CCONTCOM,");
                            sbGP.Append("TTICONPS, FBAJACON, CSEGMERC, DMUNICIP, DPROVINC, CNIFDNIC,");
                            sbGP.Append("DAPERSOC, CUPSREE, TPERFCNS, TFACTURA, TPUNTMED, TINDTUR,");
                            sbGP.Append("TEQMEDIN, PROVMUNI, CLASESUM, TDIHORAA, CODFINCA, TELEMEDIDA) values ");
                            firstOnlyGP = false;
                        }

                        sbGP.Append("('").Append(campos[i].Trim()).Append("', "); i++; // EMPRESA
                        sbGP.Append("'").Append(campos[i].Trim()).Append("', "); i++; // IDU
                        sbGP.Append("'").Append(campos[i].Trim().Replace("'", "´")).Append("', "); i++; // DDISTRIB
                        sbGP.Append(campos[i].Trim()).Append(", "); i++; // CCONTATR
                        sbGP.Append(campos[i].Trim()).Append(", "); i++; // CNUMCATR

                        for (int x = 0; x <= 9; x++)
                        {
                            sbGP.Append("'").Append(campos[i].Trim()).Append("', "); i++;
                        }

                        for (int x = 0; x <= 6; x++)
                        {
                            sbGP.Append(campos[i].Trim()).Append(", "); i++;
                        }

                        for (int x = 0; x <= 11; x++)
                        {
                            sbGP.Append("'").Append(campos[i].Trim().Replace("'", "´")).Append("', "); i++;
                        }

                        sbGP.Append(campos[i].Trim()).Append(", "); i++; // TPUNTMED

                        for (int x = 0; x <= 5; x++)
                        {
                            sbGP.Append("'").Append(campos[i].Trim()).Append("', "); i++;
                        }

                        sbGP.Append("'").Append(campos[i].Trim()).Append("'),");

                        if (numLineaGP == maxLineas)
                        {
                            Console.WriteLine("Leyendo " + totalLineas + " lineas...");
                            firstOnlyGP = true;
                            db = new MySQLDB(MySQLDB.Esquemas.CON);
                            command = new MySqlCommand(sbGP.ToString().Substring(0, sbGP.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sbGP = null;
                            sbGP = new StringBuilder();
                            numLineaGP = 0;
                        }
                    }

                } // while

                file.Close();

                if (numLineaOP > 0)
                {
                    Console.WriteLine("Leyendo " + totalLineas + " lineas...");
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sbOP.ToString().Substring(0, sbOP.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sbOP = null;
                    sbOP = new StringBuilder();
                }

                if (numLineaGP > 0)
                {
                    Console.WriteLine("Leyendo " + totalLineas + " lineas...");
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sbGP.ToString().Substring(0, sbGP.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sbGP = null;
                    sbGP = new StringBuilder();

                }
            }



            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public  bool CargaTablaPS_GP(string archivo)
        {
            EndesaBusiness.utilidades.Global global = new EndesaBusiness.utilidades.Global();
            String strSql = null;
            MySQLDB db;
            MySqlCommand command;
            DateTime inicio = new DateTime();
            DateTime fin = new DateTime();
            try
            {
                // Borramos la tabla temporal
                inicio = DateTime.Now;
                strSql = "DELETE FROM ps_total_temp;";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                fin = DateTime.Now;
                global.SaveProcess("PS", strSql + " Procesado correctamente", inicio, inicio, fin);

                // Cargamos registros en ps_total_temp
                //inicio = DateTime.Now;
                //strSql = "LOAD DATA LOCAL INFILE '" + archivo.Replace(@"\", "\\\\")
                //+ "' REPLACE INTO TABLE ps_total_temp FIELDS TERMINATED BY '|' LINES TERMINATED BY '\\n';";
                //Console.WriteLine(strSql);
                //ficheroLog.Add(strSql);

                //db = new MySQLDB(MySQLDB.Esquemas.CON);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();
                //fin = DateTime.Now;
                //global.SaveProcess("PS", strSql + " Procesado correctamente", inicio, inicio, fin);

                CargaPorLinea_PS_GP(archivo);

                // Arreglamos el campo NIF
                //inicio = DateTime.Now;
                //strSql = "UPDATE ps_total_temp SET CNIFDNIC = Trim(CNIFDNIC)";
                //Console.WriteLine(strSql);
                //db = new MySQLDB(MySQLDB.Esquemas.CON);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();
                //fin = DateTime.Now;
                //global.SaveProcess("PS", strSql + " Procesado correctamente", inicio, inicio, fin);

                // Borramos registros pertenecientes a GP
                inicio = DateTime.Now;
                strSql = "DELETE FROM ps_total WHERE CSEGMERC <> 'N';";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                fin = DateTime.Now;
                global.SaveProcess("PS", strSql + " Procesado correctamente", inicio, inicio, fin);

                // Anexamos la tabla temporal
                inicio = DateTime.Now;
                strSql = "REPLACE INTO ps_total SELECT * FROM ps_total_temp;";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                fin = DateTime.Now;
                global.SaveProcess("PS", strSql + " Procesado correctamente", inicio, inicio, fin);

                return true;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                fin = DateTime.Now;
                global.SaveProcess("PS", "ERROR: " + utilidades.FuncionesTexto.ArreglaAcentos(e.Message), inicio, inicio, fin);
                return false;
            }

        }

        public bool CargaTablaPS_OP(string archivo)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            try
            {
                ss_pp.Update_Fecha_Inicio("Contratación", "PS_AT", "Importación PS");

                ficheroLog.Add("CargaTablaPS_OP");
                ficheroLog.Add("===============");
                ficheroLog.Add("");
                strSql = "DELETE FROM ps_total WHERE CSEGMERC = 'N'";
                ficheroLog.Add("INICIO: " + strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                ficheroLog.Add("FIN: " + strSql);
                

                strSql = "LOAD DATA LOCAL INFILE '" + archivo.Replace(@"\", "\\\\")
                    + "' REPLACE INTO TABLE ps_total"
                    + " FIELDS TERMINATED BY '|' LINES TERMINATED BY '\\n' ";
                ficheroLog.Add("INICIO: " + strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                ficheroLog.Add("FIN: " + strSql);

                strSql = "UPDATE ps_total SET ps_total.CNIFDNIC = Trim(CNIFDNIC) WHERE CSEGMERC = 'N'";
                ficheroLog.Add("INICIO: " + strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                ficheroLog.Add("FIN: " + strSql);

                ss_pp.Update_Fecha_Fin("Contratación", "PS_AT", "Importación PS");

                return true;
            }
            catch(Exception e)
            {
                ficheroLog.AddError("CargaTablaPS_OP: " + e.Message);
                ss_pp.Update_Comentario("Contratación", "PS_AT", "Importación PS", e.Message);
                return false;
            }
        }

        private string GetParameter(string codigo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            string vcodigo;
            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                strSql = "Select valor from ps_parametros where " +
                "codigo = '" + codigo + "';";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    vcodigo = reader["valor"].ToString();
                }
                else
                {
                    Console.WriteLine("El valor " + codigo + " no está parametrizado en fo_param.");
                    vcodigo = "";
                }
                db.CloseConnection();
                return vcodigo;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return "";

        }

        private string ListaDestinatarios(string tipo)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string destinatarios = null;
            bool firstOnly = true;

            try
            {
                strSql = "SELECT nombre FROM ps_destinatarios WHERE tipo = '" + tipo + "'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    if (r["nombre"] != System.DBNull.Value)
                    {
                        if (firstOnly)
                        {
                            destinatarios = destinatarios + r["nombre"].ToString();
                            firstOnly = false;
                        }
                        else
                            destinatarios = destinatarios + ";" + r["nombre"].ToString();
                    }
                        
                    
                }
                db.CloseConnection();
                return destinatarios;
            }
            catch(Exception e)
            {
                ficheroLog.AddError("ListaDestinatarios: " + e.Message);
                return null;
            }
        }

        public void EnviaInventarioBTE(FileInfo archivo)
        {
            //mail.MailExchangeServer email = new mail.MailExchangeServer(System.Environment.UserName);
            EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

            EndesaBusiness.utilidades.ZIP zip = new utilidades.ZIP();

            utilidades.UltimateFTP ftp = new utilidades.UltimateFTP(GetParameter("ftp_servidor"),
                GetParameter("FTP_contrata_usuario"), GetParameter("FTP_contrata_pass"),
                GetParameter("FTP_puerto2"));

            StringBuilder textBody = new StringBuilder();

            try
            {
                string archivoComprimido = archivo.FullName.Replace(".txt", ".zip");
                zip.Comprmir(archivo.FullName, archivoComprimido);
                FileInfo fichero = new FileInfo(archivoComprimido);
                ftp.Upload(GetParameter("FTP_contrata_ruta") + fichero.Name, fichero.FullName);


                textBody.Clear();
                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Ya se ha actualizado en el FTP el archivo ").Append(fichero.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");


                mes.SendMail("contratacionee@enel.com",
                                GetParameter("Mail_SCE"),
                               null,
                               "Inventario BTE " + DateTime.Now.ToString("dd/MM/yyyy"),
                              textBody.ToString(), null);

            }
            catch(Exception e)
            {
                ficheroLog.AddError("EnviaInventarioBTE: " + e.Message);
            }

        }

        public void CargaPorLinea_PS_GP(string archivo)
        {
            StringBuilder sb = new StringBuilder();
            MySQLDB db;
            MySqlCommand command;
            string line = "";
            bool firstOnly = true;
            string[] campos;
            long totalLineas = 0;
            long i = 0;
            int numLineas = 0;            

            try
            {
                System.IO.StreamReader file = new System.IO.StreamReader(archivo, System.Text.Encoding.GetEncoding(1252));
                while ((line = file.ReadLine()) != null)
                {
                    totalLineas++;
                    numLineas++;
                    campos = line.Split('|');
                    i = 0;

                    if (firstOnly)
                    {
                        sb.Append("REPLACE INTO ps_total_temp (EMPRESA, IDU, DDISTRIB, CCONTATR, CNUMCATR, FPREALTA,");
                        sb.Append("FALTACON, FPSERCON, FPREVBAJ, TESTCONT, CTARIFA, VTENSIOM,");
                        sb.Append("TDISCHOR, CONTREXT, CNUMSCPS, VPOTCAL1, VPOTCAL2, VPOTCAL3,");
                        sb.Append("VPOTCAL4, VPOTCAL5, VPOTCAL6, CONSESTI, TINDGCPY, CCONTCOM,");
                        sb.Append("TTICONPS, FBAJACON, CSEGMERC, DMUNICIP, DPROVINC, CNIFDNIC,");
                        sb.Append("DAPERSOC, CUPSREE, TPERFCNS, TFACTURA, TPUNTMED, TINDTUR,");
                        sb.Append("TEQMEDIN, PROVMUNI, CLASESUM, TDIHORAA, CODFINCA, TELEMEDIDA) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(campos[i].Trim()).Append("', "); i++; // EMPRESA
                    sb.Append("'").Append(campos[i].Trim()).Append("', "); i++; // IDU
                    sb.Append("'").Append(campos[i].Trim().Replace("'", "´")).Append("', "); i++; // DDISTRIB
                    sb.Append(campos[i].Trim()).Append(", "); i++; // CCONTATR
                    sb.Append(campos[i].Trim()).Append(", "); i++; // CNUMCATR

                    for (int x = 0; x <= 9; x++)
                    {
                        sb.Append("'").Append(campos[i].Trim()).Append("', "); i++;
                    }

                    for (int x = 0; x <= 6; x++)
                    {
                        sb.Append(campos[i].Trim()).Append(", "); i++;
                    }

                    for (int x = 0; x <= 11; x++)
                    {                        
                        sb.Append("'").Append(utilidades.FuncionesTexto.RT(campos[i])).Append("', "); i++;
                    }

                    sb.Append(campos[i].Trim()).Append(", "); i++; // TPUNTMED

                    for (int x = 0; x <= 5; x++)
                    {
                        sb.Append("'").Append(campos[i].Trim()).Append("', "); i++;
                    }

                    sb.Append("'").Append(campos[i].Trim()).Append("'),");

                    if (numLineas == 500)
                    {
                        Console.WriteLine("Leyendo " + totalLineas + " lineas...");
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        numLineas = 0;



                    }
                }
                file.Close();

                if (numLineas > 0)
                {
                    Console.WriteLine("Leyendo " + totalLineas + " lineas...");
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;                    

                }


            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }


    }
}
