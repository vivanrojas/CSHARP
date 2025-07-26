using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class StarBeat_EE
    {


        logs.Log ficheroLog;
        utilidades.Param p;

        public StarBeat_EE()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "StarBeat_Inventario");
            p = new utilidades.Param("starbeat_inventario_param", servidores.MySQLDB.Esquemas.MED);
        }

        public void Descarga()
        {
            string[] lista_archivos;
            string dirFTP = "";
            
            List<string> rutas_validas = new List<string>();
            DateTime UltimaFecha = UltimaFechaArchivoProcesado();



            ficheroLog.Add("Conectando a FTP: " + p.GetValue("ftp_server", DateTime.Now, DateTime.Now));
            utilidades.FTP ftpClient = new utilidades.FTP(p.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                    p.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                    p.GetValue("ftp_pass", DateTime.Now, DateTime.Now));


            DirectoryInfo rutaDescarga = new DirectoryInfo(p.GetValue("inbox", DateTime.Now, DateTime.Now));
            if (!rutaDescarga.Exists)
                rutaDescarga.Create();

            dirFTP = p.GetValue("ftp_ruta_starbeat", DateTime.Now, DateTime.Now);
            lista_archivos = ftpClient.DirectoryListSimple(dirFTP);
            for (int i = 0; i < lista_archivos[i].Count(); i++)
            {
                if (lista_archivos[i].Contains(p.GetValue("prefijo_starbeat", DateTime.Now, DateTime.Now))
                    && (FechaArchivo(lista_archivos[i]) > UltimaFecha))
                {
                    Console.CursorLeft = 0;
                    Console.Write("Descargando " + lista_archivos[i]);
                    ficheroLog.Add("Descargando " + lista_archivos[i]);
                    ftpClient.Download(@"/" + dirFTP + lista_archivos[i], rutaDescarga.FullName + lista_archivos[i]);
                }

            }

          

        }

        public void Carga()
        {
            string[] listaArchivosZIP;
            string[] listaArchivosTXT;
            utilidades.ZIP zip = new utilidades.ZIP();

            listaArchivosZIP = Directory.GetFiles(p.GetValue("inbox", DateTime.Now, DateTime.Now), "*.ZIP");
            for (int i = 0; i < listaArchivosZIP.Length; i++)
            {
                FileInfo fichero = new FileInfo(listaArchivosZIP[i]);
                ficheroLog.Add("Descomprimiento " + fichero.Name);
                zip.Descomprimir(listaArchivosZIP[i], p.GetValue("inbox", DateTime.Now, DateTime.Now), null);
                fichero.Delete();
            }

            listaArchivosTXT = Directory.GetFiles(p.GetValue("inbox", DateTime.Now, DateTime.Now), "*.TXT");
            for (int i = 0; i < listaArchivosTXT.Length; i++)
            {
                ficheroLog.Add("Importando " + listaArchivosTXT[i]);
                Importar(listaArchivosTXT[i]);
            }

            if (p.GetValue("enviar_mail_fin_proceso", DateTime.Now, DateTime.Now) == "S")
                MailFinProceso();
        }


        public void Importar(string fichero)
        {
            string[] c;
            int numLine = 0;
            int i = 0;
            string line = "";
            List<EndesaEntity.Kee_Inventario_StarBeat> lista = new List<EndesaEntity.Kee_Inventario_StarBeat>();
            utilidades.UtilidadesArchivo utilArch = new utilidades.UtilidadesArchivo();
            int totalLineasFichero = 0;
            bool hayError;

            try
            {
                FileInfo file = new FileInfo(fichero);
                totalLineasFichero = utilArch.LineasArchivo(file, 1);

                System.IO.StreamReader archivo = new System.IO.StreamReader(file.FullName);
                while ((line = archivo.ReadLine()) != null)
                {
                    numLine++;
                    c = line.Split(';');

                    i = 0;
                    if (numLine > 1)
                    {
                        EndesaEntity.Kee_Inventario_StarBeat f = new EndesaEntity.Kee_Inventario_StarBeat();
                        f.archivo = file.Name;
                        f.pm = c[i]; i++;
                        f.num_serie_15 = c[i]; i++;
                        f.serie_22 = c[i]; i++;
                        if (c[i] != "")                           
                        {
                            f.fecha_alta = c[i].Substring(0, 19);
                            if(c[i].Length > 20)
                                f.fecha_alta_ms = Convert.ToInt32(c[i].Substring(20, c[i].Length - 20));
                        }
                        i++;
                        f.tipo_punto_medida = Convert.ToInt32(c[i]); i++;
                        f.tipo_via_principal = c[i]; i++;
                        f.ip = c[i]; i++;
                        if(c[i] != "")
                            f.puerto = Convert.ToInt32(c[i]);
                        i++;
                                          
                        f.telefono = c[i]; i++;
                        f.de = c[i]; i++;
                        f.dpm = c[i]; i++;

                        if (c[i] != "")
                            f.ult_lect_1cc = c[i]; 
                        i++;
                        if (c[i] != "")
                            f.ult_lect_2cc = c[i]; 
                        i++;
                        if (c[i] != "")
                            f.ult_lect_atr = c[i];
                        i++;

                        f.inhibido = Convert.ToInt32(c[i]);
                        i++;

                        if (c[i] != "")
                        {
                            f.fecha_cambio = c[i].Substring(0, 19);
                            if (c[i].Length > 20)
                                f.fecha_cambio_ms = Convert.ToInt32(c[i].Substring(20, c[i].Length - 20));
                        }

                        lista.Add(f);

                    }
                }
                archivo.Close();
                
                
                hayError = GuardaDatos(lista);
                if (!hayError)
                {
                    GuardaArchivo(file.Name, totalLineasFichero, lista.Count);
                    file.Delete();
                }
                    
                
            }
            catch(Exception e)
            {
                ficheroLog.AddError("Archivo: " + fichero +  " linea: "  + numLine + " " + e.Message);
                MailError("Archivo: " + fichero + " linea: " + numLine + " " + e.Message);
            }
            
        }

        private bool GuardaDatos(List<EndesaEntity.Kee_Inventario_StarBeat> lista)
        {
            string strSql = "";
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            MySQLDB db;
            MySqlCommand command;
            int x = 0;
            int i = 0;

            try
            {

                strSql = "replace into starbeat_inventario_hist select * from starbeat_inventario";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "delete from starbeat_inventario";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                foreach (EndesaEntity.Kee_Inventario_StarBeat p in lista)
                {
                    x++;
                    i++;
                    if (firstOnly)
                    {
                        sb.Append("replace into starbeat_inventario (");
                        sb.Append("pm, num_serie_15, serie_22, fecha_alta, fecha_alta_ms, tipo_punto_medida, tipo_via_principal,");
                        sb.Append("ip, puerto, telefono, de, dpm, ult_lect_1cc, ult_lect_2cc, ult_lect_atr, inhibido, fecha_cambio,");
                        sb.Append("fecha_cambio_ms, archivo) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(p.pm).Append("',");
                    sb.Append("'").Append(p.num_serie_15).Append("',");
                    sb.Append("'").Append(p.serie_22).Append("',");


                    if (p.fecha_alta != null)
                    {
                        sb.Append("'").Append(p.fecha_alta).Append("',");
                        if (p.fecha_alta_ms != 0)
                            sb.Append(p.fecha_alta_ms).Append(",");
                        else
                            sb.Append("null,");
                    }                        
                    else
                        sb.Append("null, null,");

                    sb.Append(p.tipo_punto_medida).Append(",");

                    if (p.telefono != null)
                        sb.Append("'").Append(p.tipo_via_principal).Append("',");
                    else
                        sb.Append("null,");


                    if (p.ip != null)
                        sb.Append("'").Append(p.ip).Append("',");
                    else
                        sb.Append("null,");

                    if (p.puerto != 0)
                        sb.Append(p.puerto).Append(",");
                    else
                        sb.Append("null,");

                    if(p.telefono != null)
                        sb.Append("'").Append(p.telefono).Append("',");
                    else
                        sb.Append("null,");

                    if (p.de != null)
                        sb.Append("'").Append(p.de).Append("',");
                    else
                        sb.Append("null,");

                    if (p.dpm != null)
                        sb.Append("'").Append(p.dpm).Append("',");
                    else
                        sb.Append("null,");

                    if (p.ult_lect_1cc != null)
                        sb.Append("'").Append(p.ult_lect_1cc).Append("',");
                    else
                        sb.Append("null,");

                    if (p.ult_lect_2cc != null)
                        sb.Append("'").Append(p.ult_lect_2cc).Append("',");
                    else
                        sb.Append("null,");

                    if (p.ult_lect_atr != null)
                        sb.Append("'").Append(p.ult_lect_atr).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append(p.inhibido).Append("',");
                    sb.Append("'").Append(p.fecha_cambio).Append("',");
                    if (p.fecha_cambio_ms != 0)
                        sb.Append(p.fecha_cambio_ms).Append(",");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append(p.archivo).Append("'),");


                    if (x == 250)
                    {
                        Console.CursorLeft = 0;
                        Console.Write("Guardando " + i + " registros...");
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.MED);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        x = 0;
                    }
                }

                if (x > 0)
                {
                    Console.CursorLeft = 0;
                    Console.Write("Guardando " + i + " registros...");
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    x = 0;
                }

                return false;
            }
            catch(Exception e)
            {
                ficheroLog.Add(e.Message);
                MailError("GuardaDatos" + e.Message);
                return true;
            }
        }

        private DateTime FechaArchivo(string nombreArchivo)
        {
            string fechaArchivo = "";
            DateTime fecha = new DateTime();

            fechaArchivo =
                nombreArchivo.Replace(p.GetValue("prefijo_starbeat", DateTime.Now, DateTime.Now), "");
            fechaArchivo =
                fechaArchivo.Replace(p.GetValue("extension_starbeat", DateTime.Now, DateTime.Now), "");           

            fecha = new DateTime(Convert.ToInt32(fechaArchivo.Substring(0, 4)),
                Convert.ToInt32(fechaArchivo.Substring(4, 2)),
                Convert.ToInt32(fechaArchivo.Substring(6, 2)));

            return fecha;
        }

        private DateTime UltimaFechaArchivoProcesado()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            strSql = "select max(fecha_archivo) as ultimaFecha from starbeat_inventario_archivos";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();

            while (r.Read())
            {
                if(r["ultimaFecha"] != System.DBNull.Value)
                    return Convert.ToDateTime(r["ultimaFecha"]);
            }
            db.CloseConnection();

            return DateTime.MinValue;
        }

        private string UltimoArchivoProcesado()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            strSql = "SELECT nombre_archivo, fecha_archivo  from starbeat_inventario_archivos a"
                + " INNER JOIN(SELECT MAX(fecha_archivo) maximaFecha FROM starbeat_inventario_archivos) aa ON"
                + " aa.maximaFecha = a.fecha_archivo";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();

            while (r.Read())
            {
                if (r["nombre_archivo"] != System.DBNull.Value)
                    return r["nombre_archivo"].ToString();
            }
            db.CloseConnection();

            return "";
        }

        private void GuardaArchivo(string nombreArchivo, int lineasArchivo, int lineasProcesadas)
        {
            string strSql = "";            
            MySQLDB db;
            MySqlCommand command;

            strSql = "replace into starbeat_inventario_archivos ("
                + " nombre_archivo, fecha_archivo, lineas_archivo_sin_cabecera,"
                + " lineas_procesadas) values ("
                + "'" + nombreArchivo + "',"
                + "'" + FechaArchivo(nombreArchivo).ToString("yyyy-MM-dd") + "',"
                + lineasArchivo + ","
                + lineasProcesadas + ")";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        public void MailFinProceso()
        {

            string saludo = DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:";

            string mensaje = saludo +
               System.Environment.NewLine +
               System.Environment.NewLine +
               "Ha finalizado el proceso de carga de StarBeat EE " +
               System.Environment.NewLine +
               System.Environment.NewLine +
               "Se ha procesado el archivo " + UltimoArchivoProcesado() +
               "." +
               System.Environment.NewLine +
               System.Environment.NewLine +
               "Un saludo.";


            //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("ES02255021D");
            EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
            mes.SendMail("rsiope.gma@enel.com", 
                p.GetValue("mail_to", DateTime.Now,DateTime.Now),
                p.GetValue("mail_cc", DateTime.Now, DateTime.Now),
                "Fin carga StarBeat Inventario. " + DateTime.Now.ToString("dd/MM/yyyy"),
               mensaje, null);
        }

        public void MailError(string mensajeError)
        {
            //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("ES02255021D");
            EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
            mes.SendMail("rsiope.gma@enel.com", "rsiope.gma@enel.com", null,
                "Error en carga StarBeat EE " + DateTime.Now.ToString("dd/MM/yyyy"),
               mensajeError, null);
        }


    }
}
