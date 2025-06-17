using EndesaBusiness.servidores;
using EndesaEntity.global;
using Microsoft.Office.Core;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MySQLDB = EndesaBusiness.servidores.MySQLDB;

namespace EndesaBusiness.contratacion
{
    public class Portugal
    {
        utilidades.Param p;
        logs.Log ficheroLog;
        Dictionary<string, EndesaEntity.contratacion.Tabla_cp_fechas_procesos> dic_fechas_procesos;
        Dictionary<DateTime, List<EndesaEntity.contratacion.Tabla_cp_fechas_procesos>> dic_log_archivos_zip;
        Dictionary<DateTime, List<EndesaEntity.contratacion.Tabla_cp_fechas_procesos>> dic_log_archivos_txt;
        Dictionary<string, List<string>> dic_lista_archivos_control;
        utilidades.Seguimiento_Procesos ss_pp;


        public Portugal()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_ContratacionPortugal_SharePoint");
            p = new utilidades.Param("cp_param", servidores.MySQLDB.Esquemas.CON);
            dic_fechas_procesos = CargaFechasProcesos();
            dic_lista_archivos_control = new Dictionary<string, List<string>>();
            ss_pp = new utilidades.Seguimiento_Procesos();
        }

        private Dictionary<string, EndesaEntity.contratacion.Tabla_cp_fechas_procesos> CargaFechasProcesos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

           

            Dictionary<string, EndesaEntity.contratacion.Tabla_cp_fechas_procesos>  d =
                new Dictionary<string, EndesaEntity.contratacion.Tabla_cp_fechas_procesos>();
                
            try
            {

                strSql = "select proceso, extractor, fecha from cp_fechas_procesos";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    EndesaEntity.contratacion.Tabla_cp_fechas_procesos c =
                        new EndesaEntity.contratacion.Tabla_cp_fechas_procesos();

                    c.proceso = r["proceso"].ToString();
                    c.extractor = r["extractor"].ToString();

                    if (r["fecha"] != System.DBNull.Value)
                        c.fecha = Convert.ToDateTime(r["fecha"]);
                    else
                        c.fecha = DateTime.MinValue;

                    EndesaEntity.contratacion.Tabla_cp_fechas_procesos o;
                    if(!d.TryGetValue(c.proceso, out o))
                        d.Add(c.proceso, c);
                }
                db.CloseConnection();

                return d;
            }catch(Exception ex)
            {
                ficheroLog.AddError("CargaFechasProcesos: " + ex.Message);
                return null;
            }
        }

        private Dictionary<DateTime, List<EndesaEntity.contratacion.Tabla_cp_fechas_procesos>> CargaLogArchivos(string tabla)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime fecha = new DateTime();


            Dictionary<DateTime, List<EndesaEntity.contratacion.Tabla_cp_fechas_procesos>> d =
                new Dictionary<DateTime, List<EndesaEntity.contratacion.Tabla_cp_fechas_procesos>>();

            try
            {

                strSql = "select archivo, fecha_ejecucion, kb, md5 from " + tabla;
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    EndesaEntity.contratacion.Tabla_cp_fechas_procesos c =
                        new EndesaEntity.contratacion.Tabla_cp_fechas_procesos();

                    c.archivo = r["archivo"].ToString();
                    fecha = Convert.ToDateTime(r["fecha_ejecucion"]);
                    c.kb = Convert.ToInt32(r["kb"]);
                    c.md5 = r["md5"].ToString();
                    

                    List<EndesaEntity.contratacion.Tabla_cp_fechas_procesos> o;
                    if (!d.TryGetValue(fecha, out o))
                    {
                        o = new List<EndesaEntity.contratacion.Tabla_cp_fechas_procesos>();
                        o.Add(c);
                        d.Add(fecha, o);
                    }
                    else
                        o.Add(c);

                        
                }
                db.CloseConnection();

                return d;
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("CargaLogArchivos: " + ex.Message);
                return null;
            }
        }

        private Dictionary<string, List<string>> CargaControlArchivos()
        {

            // Carga una lista de archivos que debemos tener controlados

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string extractor = "";
            string archivo = "";

            Dictionary<string, List<string>> d =
                new Dictionary<string, List<string>>();

            try
            {
                strSql = "select archivo, extractor from cp_lista_archivos_a_controlar";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    extractor = r["extractor"].ToString();
                    archivo = r["archivo"].ToString();

                    List<string> o;
                    if (!d.TryGetValue(extractor, out o))
                    {
                        o = new List<string>();
                        o.Add(archivo);
                        d.Add(extractor, o);
                    }
                    else
                        o.Add(archivo);
                }
                db.CloseConnection();



            }
            catch(Exception ex)
            {
                ficheroLog.AddError("CargaControlArchivos: " + ex.Message);
                return null;
            }

            return d;
        }

        private bool ExistenArchivosExtractor(string extractor, DateTime fecha,
            Dictionary<string, List<string>> dic_lista_archivos_control,
            Dictionary<DateTime, List<EndesaEntity.contratacion.Tabla_cp_fechas_procesos>> dic_log_archivos_zip)
        {
            bool existe = false;

            List<EndesaEntity.contratacion.Tabla_cp_fechas_procesos> lista_archivos;
            if (dic_log_archivos_zip.TryGetValue(fecha, out lista_archivos))
            {

            }
            else
                return false;

            List<string> lista_archivos_control;
            if(dic_lista_archivos_control.TryGetValue(extractor, out lista_archivos_control))
            {
                foreach(string p in lista_archivos_control)
                {
                    if(lista_archivos.Exists(z => z.archivo.Contains(p)))
                    {
                        existe = true;
                    }
                    else
                    {
                        ficheroLog.Add("No existe el archivo: " + p);
                        return false;
                    }
                }
                    
                    
                
            }

            return existe;
        }






        public void EnviarArchivosFecha()
        {
            EndesaBusiness.utilidades.Global global = new EndesaBusiness.utilidades.Global();
            //String ultimoDiaHabil = utilidades.Fichero.UltimoDiaHabil_YYMMDD();
            string ultimoDiaHabil = p.GetValue("Fecha_Envio_Archivos");
            FileInfo file;
            DateTime inicio = new DateTime();
            string ficheroGordo = "";
            int desdeLinea = 0;
            int desdeColumna = 0;
            int hastaColumna = 0;

            utilidades.UltimateFTP ftp;
            utilidades.ZipUnZip zip;

            DirectoryInfo dirSalida;

            try
            {
                zip = new utilidades.ZipUnZip();

                ficheroGordo = p.GetValue("strSOLATRBTN_YYMMDD");
                desdeLinea = Convert.ToInt32(p.GetValue("divideFichero_desdeLinea"));
                desdeColumna = Convert.ToInt32(p.GetValue("divideFichero_desdeColumna"));
                hastaColumna = Convert.ToInt32(p.GetValue("divideFichero_hastaColumna"));

                dirSalida = new DirectoryInfo(p.GetValue("Inbox"));
                if (!dirSalida.Exists)
                    dirSalida.Create();



                inicio = DateTime.Now;
                ficheroLog.Add("Ejecutando extractor: " + p.GetValue("ExtractorBPO_AAMMDD_F") + " " + ultimoDiaHabil);
                utilidades.Fichero.EjecutaComando(p.GetValue("ExtractorBPO_AAMMDD_F"), ultimoDiaHabil);
                global.SaveProcess("ContratacionPortugal", "ExtractorBPO_AAMMDD_F", inicio, inicio, DateTime.Now);

                inicio = DateTime.Now;
                ficheroLog.Add("Ejecutando extractor: " + p.GetValue("ExtractorBPO_MMDD_F") + ultimoDiaHabil.Substring(2, 4));
                utilidades.Fichero.EjecutaComando(p.GetValue("ExtractorBPO_MMDD_F"), ultimoDiaHabil.Substring(2, 4));
                global.SaveProcess("ContratacionPortugal", "ExtractorBPO_MMDD_F", inicio, inicio, DateTime.Now);

                inicio = DateTime.Now;
                ficheroLog.Add("Ejecutando extractor: " + p.GetValue("Extractor_BTN_AAMMDD_F") + " " + ultimoDiaHabil);
                utilidades.Fichero.EjecutaComando(p.GetValue("Extractor_BTN_AAMMDD_F"), ultimoDiaHabil);
                global.SaveProcess("ContratacionPortugal", "Extractor_BTN_AAMMDD_F", inicio, inicio, DateTime.Now);

                string[] listaArchivos = Directory.GetFiles(p.GetValue("Inbox_F"));


                ftp = new utilidades.UltimateFTP(
                       p.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                       p.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                       utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_pass", DateTime.Now, DateTime.Now), true),
                       p.GetValue("ftp_port", DateTime.Now, DateTime.Now));


                if (listaArchivos.Length > 0)
                {


                    for (int i = 0; i < listaArchivos.Length; i++)
                    {
                        inicio = DateTime.Now;
                        file = new FileInfo(listaArchivos[i]);

                        if (listaArchivos[i].Contains(ficheroGordo))
                            TrataFicheroGordo(listaArchivos[i], desdeLinea, desdeColumna, hastaColumna);

                        zip.ComprimirArchivo(listaArchivos[i], listaArchivos[i].Replace(".txt", ".zip"));
                        global.SaveProcess("ContratacionPortugal", "Comprimiendo " + file.Name, inicio, inicio, DateTime.Now);
                        file.Delete();

                    }
                }
                else
                {

                }
                                

                listaArchivos = Directory.GetFiles(p.GetValue("Inbox_F"), "*.zip");

                for (int i = 0; i < listaArchivos.Length; i++)
                {
                    FileInfo fichero = new FileInfo(listaArchivos[i]);
                    ftp.Upload(p.GetValue("ruta_destino_FTP") + fichero.Name, listaArchivos[i]);
                    utilidades.Fichero.BorrarArchivo(listaArchivos[i]);
                }

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Portugal.EnviarArchivos: " + e.Message);
            }
        }

        

        public void LanzaTodo()
        {
            EjecutaProceso("BPO-ContratacionAAMMDD", "ExtractorBPO_AAMMDD");
            EjecutaProceso("BPO-Contratacion-MMDD", "ExtractorBPO_MMDD");
            EjecutaProceso("BPO-Contratacion_SOLBTN_AAMMDD", "Extractor_BTN_AAMMDD");
        }

        private Dictionary<string, EndesaEntity.herramientas.Seguimiento_Procesos> Carga_Control_Archivos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            Dictionary<string, EndesaEntity.herramientas.Seguimiento_Procesos> d =
                new Dictionary<string, EndesaEntity.herramientas.Seguimiento_Procesos>();

            try
            {

                strSql = "SELECT ca.prefijo_archivo, ca.extractor, ca.paso, formato_fecha"
                    + " FROM cp_control_archivos ca"
                    + " WHERE ca.extractor IS NOT NULL AND ca.paso IS NOT null";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.herramientas.Seguimiento_Procesos c = new EndesaEntity.herramientas.Seguimiento_Procesos();
                    c.prefijo_archivo = r["prefijo_archivo"].ToString();
                    c.extractor = r["extractor"].ToString();
                    c.descripcion = r["paso"].ToString();
                    c.comentario = r["formato_fecha"].ToString();
                    d.Add(c.prefijo_archivo, c);
                }
                db.CloseConnection();
                return d;
            }
            catch(Exception ex)
            {
                return null;
            }
        }


        private Dictionary<string, string> ArchivosZIP_Ultimos15Dias()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string archivo = "";

            Dictionary<string, string> d = new Dictionary<string, string>();

            try
            {

                strSql = "select archivo from cp_log_archivos_zip"
                    + " where fecha_ejecucion > '" + DateTime.Now.AddDays(-15).Date.ToString("yyyy-MM-dd") + "'"
                    + " order by fecha_ejecucion";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    archivo = r["archivo"].ToString();
                    d.Add(archivo, archivo);
                }
                db.CloseConnection();
                return d;

            }
            catch(Exception ex)
            {
                return null;
            }
        }

        public void Nuevo_Proceso_Archivos_Portugal()
        {
            DateTime fecha = new DateTime();
            Dictionary<string, string> dic_log_archivos_zip;
            Dictionary<string, EndesaEntity.herramientas.Seguimiento_Procesos> dic_archivo_extractor;

            string fecha_string = "";
            string archivo = "";
            FileInfo fileinfo;
            FileInfo archivo_zip;
            FileInfo archivo_destino;
            string md5 = "";
            utilidades.ZipUnZip zip7z = new utilidades.ZipUnZip();

            try
            {
                #region Achivos-Extractores
                dic_archivo_extractor = Carga_Control_Archivos();
               
                #endregion



                utilidades.Fechas utilFechas = new utilidades.Fechas();

                fecha = utilFechas.UltimoDiaHabil();
                dic_log_archivos_zip = ArchivosZIP_Ultimos15Dias();


                for (DateTime d = fecha.Date; d > fecha.Date.AddDays(-15); d = utilFechas.UltimoDiaHabilAnterior(d))
                {
                    Console.WriteLine(d.ToString("yyyy-MM-dd"));

                    foreach (KeyValuePair<string, EndesaEntity.herramientas.Seguimiento_Procesos> pp in dic_archivo_extractor)
                    {

                        switch (pp.Value.comentario)
                        {
                            case "MMDD":
                                fecha_string = d.ToString("MMdd");
                                break;
                            case "AAMMDD":
                                fecha_string = d.ToString("yyMMdd");
                                break;
                        }


                        archivo = pp.Value.prefijo_archivo + fecha_string + ".txt";
                        ficheroLog.Add("archivo: " + archivo);

                        archivo_zip = new FileInfo(p.GetValue("Inbox") + "\\" + archivo.Replace(".txt", ".zip"));
                        ficheroLog.Add("archivo_zip: " + archivo_zip.FullName);

                        archivo_destino = new FileInfo(p.GetValue("ruta_sharepoint")
                            + archivo_zip.Name.Substring(0, archivo_zip.Name.IndexOf("_d"))
                            + @"\" + archivo_zip.Name);

                        ficheroLog.Add(archivo);

                        string o;
                        if (!dic_log_archivos_zip.TryGetValue(archivo.Replace(".txt", ".zip"), out o))
                        {
                            if (d.Date == utilFechas.UltimoDiaHabil().Date)
                                ss_pp.Update_Fecha_Inicio("Contratación", "Archivos Portugal SharePoint", pp.Value.descripcion);

                            Console.WriteLine("Ejecutando comando: " + pp.Value.extractor + " " + fecha_string);
                            ficheroLog.Add("Ejecutando comando: " + pp.Value.extractor + " " + fecha_string);
                            utilidades.Fichero.EjecutaComando(pp.Value.extractor, fecha_string);

                            fileinfo = new FileInfo(p.GetValue("Inbox") + "\\" + archivo);
                            if (fileinfo.Exists)
                            {
                                if (fileinfo.Length > 0)
                                {
                                    if ((fileinfo.Length / 1024) > 0)
                                    {
                                        zip7z.ComprimirArchivo(fileinfo.FullName, fileinfo.FullName.Replace(".txt", ".zip"));

                                        if (archivo_destino.Exists)
                                            archivo_destino.Delete();

                                        archivo_zip.CopyTo(archivo_destino.FullName);
                                        md5 = utilidades.Fichero.checkMD5(fileinfo.FullName).ToString();


                                        ficheroLog.Add("Actualizando proceso cp_log_archivos_zip: " + archivo_zip.Name);
                                        ActualizaArchivoFechaProceso("cp_log_archivos_zip", archivo_zip.Name, DateTime.Now, archivo_zip.Length / 1024, md5);

                                        if (d.Date == utilFechas.UltimoDiaHabil().Date)
                                        {
                                            ss_pp.Update_Fecha_Fin("Contratación", "Archivos Portugal SharePoint", pp.Value.descripcion);
                                        }
                                    }
                                    else
                                    {
                                        if (d.Date == utilFechas.UltimoDiaHabil().Date)
                                        {
                                            ss_pp.Update_Comentario("Contratación", "Archivos Portugal SharePoint",
                                                    pp.Value.descripcion, "No existe el archivo " + fileinfo.Name);


                                        }

                                        ficheroLog.Add("No existe el archivo " + fileinfo.Name);
                                    }
                                }
                                else
                                {
                                    if (d.Date == utilFechas.UltimoDiaHabil().Date)
                                        ss_pp.Update_Comentario("Contratación", "Archivos Portugal SharePoint",
                                                pp.Value.descripcion, "No existe el archivo " + fileinfo.Name);

                                    ficheroLog.Add("No existe el archivo " + fileinfo.Name);
                                }



                            }
                            else
                            {
                                if (d.Date == utilFechas.UltimoDiaHabil().Date)
                                    ss_pp.Update_Comentario("Contratación", "Archivos Portugal SharePoint",
                                            pp.Value.descripcion, "No existe el archivo " + fileinfo.Name);

                                ficheroLog.Add("No existe el archivo " + fileinfo.Name);
                            }


                            if (archivo_zip.Exists)
                            {
                                ficheroLog.Add("Borrando: " + archivo_zip.Name);
                                archivo_zip.Delete();
                            }


                            if (fileinfo.Exists)
                            {
                                ficheroLog.Add("Borrando: " + fileinfo.Name);
                                fileinfo.Delete();
                            }



                        }
                        else
                        {
                            ficheroLog.Add("El archivo: " + archivo.Replace(".txt", ".zip")
                                + " ya está en SharePoint.");
                        }

                    }
                }



            }catch(Exception ex)
            {
                ficheroLog.AddError(ex.Message);
            }

              


            
        }


        public void Descarga_Ficheros_Contratacion_Portugal(DateTime fecha_proceso)
        {
            string ultimoDiaHabil = "";
            DateTime fecha = new DateTime();
            DirectoryInfo dirSalida;
            FileInfo file;
            string md5;
            utilidades.ZipUnZip zip7z = new utilidades.ZipUnZip();
            utilidades.UltimateFTP ftp;
            string log;

            FileInfo archivo_destino;
            FileInfo archivo_zip;

            try
            {
                

                dic_log_archivos_zip = CargaLogArchivos("cp_log_archivos_zip");
                dic_lista_archivos_control = CargaControlArchivos();







                if(fecha_proceso.Date == DateTime.Now.Date)
                {
                    fecha = DateTime.Now.Date;
                    ultimoDiaHabil = utilidades.Fichero.UltimoDiaHabil_YYMMDD();
                    Console.WriteLine("Último día Hábil: " + ultimoDiaHabil);
                }
                else
                {
                    ultimoDiaHabil = p.GetValue("Fecha_Envio_Archivos");
                }
                
                

                if(!ExistenArchivosExtractor("BPO-ContratacionAAMMDD_TS.bat", fecha, dic_lista_archivos_control, dic_log_archivos_zip))
                {
                    ficheroLog.Add("Ejecutando extractor: " + p.GetValue("ExtractorBPO_AAMMDD") + " " + ultimoDiaHabil);
                    Console.WriteLine("Ejecutando extractor: " + p.GetValue("ExtractorBPO_AAMMDD") + " " + ultimoDiaHabil);
                    utilidades.Fichero.EjecutaComando(p.GetValue("ExtractorBPO_AAMMDD"), ultimoDiaHabil);
                }
                else
                {
                    log = "No se ejecuta el extractor: " + p.GetValue("ExtractorBPO_AAMMDD") + " " + ultimoDiaHabil
                        + " porque están todos los archivos procesados.";
                    ficheroLog.Add(log);
                    Console.WriteLine(log);
                }

                if (!ExistenArchivosExtractor("BPO-Contratacion-MMDD_TS.bat", fecha, dic_lista_archivos_control, dic_log_archivos_zip))
                {
                    ficheroLog.Add("Ejecutando extractor: " + p.GetValue("ExtractorBPO_MMDD") + " " + ultimoDiaHabil.Substring(2, 4));
                    Console.WriteLine("Ejecutando extractor: " + p.GetValue("ExtractorBPO_MMDD") + " " + ultimoDiaHabil.Substring(2, 4));
                    utilidades.Fichero.EjecutaComando(p.GetValue("ExtractorBPO_MMDD"), ultimoDiaHabil.Substring(2, 4));
                }
                else
                {
                    log = "No se ejecuta el extractor: " + p.GetValue("ExtractorBPO_MMDD") + " " + ultimoDiaHabil
                        + " porque están todos los archivos procesados.";
                    ficheroLog.Add(log);
                    Console.WriteLine(log);
                }


                if (!ExistenArchivosExtractor("BPO-Contratacion_SOLATRDPPARTIDO_AAMMDD_TS.bat", fecha, dic_lista_archivos_control, dic_log_archivos_zip))
                {
                    ficheroLog.Add("Ejecutando extractor: " + p.GetValue("Extractor_BTN_PARTIDO_AAMMDD") + " " + ultimoDiaHabil);
                    Console.WriteLine("Ejecutando extractor: " + p.GetValue("Extractor_BTN_PARTIDO_AAMMDD") + " " + ultimoDiaHabil);
                    utilidades.Fichero.EjecutaComando(p.GetValue("Extractor_BTN_PARTIDO_AAMMDD"), ultimoDiaHabil);
                }
                else
                {
                    log = "No se ejecuta el extractor: " + p.GetValue("Extractor_BTN_PARTIDO_AAMMDD") + " " + ultimoDiaHabil
                        + " porque están todos los archivos procesados.";
                    ficheroLog.Add(log);
                    Console.WriteLine(log);
                }

                if (!ExistenArchivosExtractor("BPO-Contratacion_SOLBTN_AAMMDD_TS.bat", fecha, dic_lista_archivos_control, dic_log_archivos_zip))
                {
                    ficheroLog.Add("Ejecutando extractor: " + p.GetValue("Extractor_BTN_AAMMDD") + " " + ultimoDiaHabil);
                    Console.WriteLine("Ejecutando extractor: " + p.GetValue("Extractor_BTN_AAMMDD") + " " + ultimoDiaHabil);
                    utilidades.Fichero.EjecutaComando(p.GetValue("Extractor_BTN_AAMMDD"), ultimoDiaHabil);
                }
                else
                {
                    log = "No se ejecuta el extractor: " + p.GetValue("Extractor_BTN_AAMMDD") + " " + ultimoDiaHabil
                        + " porque están todos los archivos procesados.";
                    ficheroLog.Add(log);
                    Console.WriteLine(log);
                }


                


                dirSalida = new DirectoryInfo(p.GetValue("Inbox"));
                if (!dirSalida.Exists)
                    dirSalida.Create();

                string[] listaArchivos = Directory.GetFiles(p.GetValue("Inbox"), "*.txt");
                for (int i = 0; i < listaArchivos.Length; i++)
                {
                    file = new FileInfo(listaArchivos[i]);

                    if((file.Length / 1024) > 5)
                    {
                        md5 = utilidades.Fichero.checkMD5(file.FullName).ToString();

                        ActualizaArchivoFechaProceso("cp_log_archivos", file.Name, DateTime.Now, file.Length / 1024, md5);

                        zip7z.ComprimirArchivo(listaArchivos[i], listaArchivos[i].Replace(".txt", ".zip"));

                        archivo_zip = new FileInfo(listaArchivos[i].Replace(".txt", ".zip"));
                        archivo_destino = new FileInfo(p.GetValue("ruta_sharepoint") 
                            + archivo_zip.Name.Substring(0, archivo_zip.Name.IndexOf("_d")) 
                            + @"\" + archivo_zip.Name);

                        if (archivo_destino.Exists)
                            archivo_destino.Delete();

                        archivo_zip.CopyTo(archivo_destino.FullName);

                        file.Delete();
                    }
                   
                }

                //ftp = new utilidades.UltimateFTP(
                //     p.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                //     p.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                //     utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_pass", DateTime.Now, DateTime.Now), true),
                //     p.GetValue("ftp_port", DateTime.Now, DateTime.Now));


                //listaArchivos = Directory.GetFiles(p.GetValue("Inbox"), "*.zip");
                //for (int i = 0; i < listaArchivos.Length; i++)
                //{                   
                //    file = new FileInfo(listaArchivos[i]);
                //    md5 = utilidades.Fichero.checkMD5(file.FullName).ToString();
                //    ActualizaArchivoFechaProceso("cp_log_archivos_zip", file.Name, DateTime.Now, file.Length / 1024, md5);
                //    ftp.Upload(p.GetValue("ruta_destino_FTP") + file.Name, listaArchivos[i]);
                //    file.Delete();
                //}

                listaArchivos = Directory.GetFiles(p.GetValue("Inbox"), "*.zip");
                for (int i = 0; i < listaArchivos.Length; i++)
                {
                    file = new FileInfo(listaArchivos[i]);
                    md5 = utilidades.Fichero.checkMD5(file.FullName).ToString();
                    ActualizaArchivoFechaProceso("cp_log_archivos_zip", file.Name, DateTime.Now, file.Length / 1024, md5);                   
                    file.Delete();
                }


                //if (fecha_proceso == DateTime.Now.Date)
                //    CreaInformeMail();

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Descarga_Ficheros_Contratacion_Portugal: " + e.Message);
            }


        }

        
        

        public void CreaInformeMail()
        {
            string body = "";
            string subject = "";
            string from = "";
            string to = "";
            string cc = null;
            string attachment = null;
            string mensaje_correo = "";

            string ultimoDiaHabil = "";

            StringBuilder sb = new StringBuilder();

            bool existe = false;

            dic_log_archivos_zip = CargaLogArchivos("cp_log_archivos_zip");
            dic_log_archivos_txt = CargaLogArchivos("cp_log_archivos");
            dic_lista_archivos_control = CargaControlArchivos();

            ultimoDiaHabil = utilidades.Fichero.UltimoDiaHabil_YYMMDD();

            foreach (KeyValuePair<string, List<string>> p in dic_lista_archivos_control)
            {

                //if (p.Key.Contains("AAMMDD"))
                //    ultimoDiaHabil = utilidades.Fichero.UltimoDiaHabil_YYMMDD();
                //else
                //    ultimoDiaHabil = ultimoDiaHabil.Substring(2, 4);

                //List<EndesaEntity.contratacion.Tabla_cp_fechas_procesos> lista_archivos;
                //if (dic_log_archivos_txt.TryGetValue(DateTime.Now.Date, out lista_archivos))
                //{
                //    if (lista_archivos.Exists(z => z.archivo.Contains(p)))
                //    {
                //        existe = true;
                //    }
                //    else
                //    {
                //        ficheroLog.Add("No existe el archivo: " + p);
                //        existe = false;
                //    }
                //}
            }


            from = p.GetValue("mail_remitente", DateTime.Now, DateTime.Now);
            to = p.GetValue("mail_to", DateTime.Now, DateTime.Now);
            cc = p.GetValue("mail_cc", DateTime.Now, DateTime.Now);
            //body = p.GetValue("html_body", DateTime.Now, DateTime.Now);
            body = mensaje_correo;
            subject = p.GetValue("mail_subject", DateTime.Now, DateTime.Now);
                        
            EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
            mes.SendMail(from, to, cc, subject, body, attachment);
        }

        private void EjecutaProceso(string proceso, string extractor)
        {
            EndesaBusiness.utilidades.Global global = new EndesaBusiness.utilidades.Global();
            String ultimoDiaHabil = utilidades.Fichero.UltimoDiaHabil_YYMMDD();
            FileInfo file;
            DateTime inicio = new DateTime();
            string ficheroGordo = "";
            int desdeLinea = 0;
            int desdeColumna = 0;
            int hastaColumna = 0;

            utilidades.UltimateFTP ftp;
            utilidades.ZIP zip;

            DirectoryInfo dirSalida;

            try
            {
                zip = new utilidades.ZIP();

                if (!ProcesoCompleto(proceso))
                {
                    dirSalida = new DirectoryInfo(p.GetValue("Inbox"));
                    if (!dirSalida.Exists)
                        dirSalida.Create();

                    if (proceso == "BPO-Contratacion-MMDD ")
                        ultimoDiaHabil = ultimoDiaHabil.Substring(2, 4);

                    inicio = DateTime.Now;
                    ficheroLog.Add("Ejecutando extractor: " + p.GetValue(extractor) + " " + ultimoDiaHabil);
                    utilidades.Fichero.EjecutaComando(p.GetValue(extractor), ultimoDiaHabil);
                    global.SaveProcess("ContratacionPortugal", extractor, inicio, inicio, DateTime.Now);                   

                    string[] listaArchivos = Directory.GetFiles(p.GetValue("Inbox"));

                    ftp = new utilidades.UltimateFTP(
                      p.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                      p.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                      utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_pass", DateTime.Now, DateTime.Now), true),
                      p.GetValue("ftp_port", DateTime.Now, DateTime.Now));

                    if (listaArchivos.Length > 0)
                    {
                        Thread.Sleep(1000);

                        for (int i = 0; i < listaArchivos.Length; i++)
                        {
                            inicio = DateTime.Now;
                            file = new FileInfo(listaArchivos[i]);

                            ActualizaFechaProceso(GetProceso(file.Name), file.Length / 1024);

                            //if (listaArchivos[i].Contains(ficheroGordo))
                            //    TrataFicheroGordo(listaArchivos[i], desdeLinea, desdeColumna, hastaColumna);

                            
                            zip.Comprmir(listaArchivos[i], listaArchivos[i].Replace(".txt", ".zip"));
                            global.SaveProcess("ContratacionPortugal", "Comprimiendo " + file.Name, inicio, inicio, DateTime.Now);
                            
                            //file.Delete();

                        }

                        listaArchivos = Directory.GetFiles(p.GetValue("Inbox"), "*.txt");
                        for (int i = 0; i < listaArchivos.Length; i++)
                        {
                            file = new FileInfo(listaArchivos[i]);
                            file.Delete();
                        }



                    }

                    listaArchivos = Directory.GetFiles(p.GetValue("Inbox"), "*.zip");

                    for (int i = 0; i < listaArchivos.Length; i++)
                    {
                        FileInfo fichero = new FileInfo(listaArchivos[i]);
                        ftp.Upload(p.GetValue("ruta_destino_FTP") + fichero.Name, listaArchivos[i]);
                        utilidades.Fichero.BorrarArchivo(listaArchivos[i]);
                    }


                }

            }
            catch (Exception e)
            {

                ficheroLog.AddError("EjecutaProceso: " + e.Message);
            }
        }

        public void EnviarArchivos()
        {
            EndesaBusiness.utilidades.Global global = new EndesaBusiness.utilidades.Global();
            String ultimoDiaHabil = utilidades.Fichero.UltimoDiaHabil_YYMMDD();
            FileInfo file;
            DateTime inicio = new DateTime();
            string ficheroGordo = "";
            int desdeLinea = 0;
            int desdeColumna = 0;
            int hastaColumna = 0;

            utilidades.UltimateFTP ftp;
            utilidades.ZipUnZip zip;

            DirectoryInfo dirSalida;

            try
            {
                zip = new utilidades.ZipUnZip();

                ficheroGordo = p.GetValue("strSOLATRBTN_YYMMDD");
                desdeLinea = Convert.ToInt32(p.GetValue("divideFichero_desdeLinea"));
                desdeColumna = Convert.ToInt32(p.GetValue("divideFichero_desdeColumna"));
                hastaColumna = Convert.ToInt32(p.GetValue("divideFichero_hastaColumna"));

                dirSalida = new DirectoryInfo(p.GetValue("Inbox"));
                if (!dirSalida.Exists)
                    dirSalida.Create();

                if (!ProcesoCompleto("BPO-ContratacionAAMMDD"))
                {
                    inicio = DateTime.Now;
                    ficheroLog.Add("Ejecutando extractor: " + p.GetValue("ExtractorBPO_AAMMDD") + " " + ultimoDiaHabil);
                    utilidades.Fichero.EjecutaComando(p.GetValue("ExtractorBPO_AAMMDD"), ultimoDiaHabil);
                    global.SaveProcess("ContratacionPortugal", "ExtractorBPO_AAMMDD", inicio, inicio, DateTime.Now);

                    string[] listaArchivos = Directory.GetFiles(p.GetValue("Inbox"));

                    ftp = new utilidades.UltimateFTP(
                      p.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                      p.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                      utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_pass", DateTime.Now, DateTime.Now), true),
                      p.GetValue("ftp_port", DateTime.Now, DateTime.Now));

                    if (listaArchivos.Length > 0)
                    {

                        for (int i = 0; i < listaArchivos.Length; i++)
                        {
                            inicio = DateTime.Now;
                            file = new FileInfo(listaArchivos[i]);



                            ActualizaFechaProceso(GetProceso(file.Name), file.Length / 1024);

                            if (listaArchivos[i].Contains(ficheroGordo))
                                TrataFicheroGordo(listaArchivos[i], desdeLinea, desdeColumna, hastaColumna);

                            zip.ComprimirArchivo(listaArchivos[i], listaArchivos[i].Replace(".txt", ".zip"));
                            global.SaveProcess("ContratacionPortugal", "Comprimiendo " + file.Name, inicio, inicio, DateTime.Now);


                            file.Delete();

                        }
                    }
                    else
                    {
                        string from = p.GetValue("Buzon_envio_email");
                        string to = p.GetValue("supervisor_proceso");
                        string cc = null;
                        string subject = "Contratación Portugal (archivos normales) " + DateTime.Now.ToString("dd/MM/yyyy");
                        string body = "No se han encontrado ficheros en el extractor: "
                            + p.GetValue("Extractor_BTN_AAMMDD") + " " + ultimoDiaHabil;

                        //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                        EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
                        mes.SendMail(from, to, cc, subject, body);
                    }

                    listaArchivos = Directory.GetFiles(p.GetValue("Inbox"), "*.zip");

                    for (int i = 0; i < listaArchivos.Length; i++)
                    {
                        FileInfo fichero = new FileInfo(listaArchivos[i]);
                        ftp.Upload(p.GetValue("ruta_destino_FTP") + fichero.Name, listaArchivos[i]);
                        utilidades.Fichero.BorrarArchivo(listaArchivos[i]);
                    }

                }

                if (!ProcesoCompleto("BPO-Contratacion-MMDD"))
                {
                    inicio = DateTime.Now;
                    ficheroLog.Add("Ejecutando extractor: " + p.GetValue("ExtractorBPO_MMDD") + ultimoDiaHabil.Substring(2, 4));
                    utilidades.Fichero.EjecutaComando(p.GetValue("ExtractorBPO_MMDD"), ultimoDiaHabil.Substring(2, 4));
                    global.SaveProcess("ContratacionPortugal", "ExtractorBPO_MMDD", inicio, inicio, DateTime.Now);

                    string[] listaArchivos = Directory.GetFiles(p.GetValue("Inbox"));

                    ftp = new utilidades.UltimateFTP(
                      p.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                      p.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                      utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_pass", DateTime.Now, DateTime.Now), true),
                      p.GetValue("ftp_port", DateTime.Now, DateTime.Now));

                    if (listaArchivos.Length > 0)
                    {


                        for (int i = 0; i < listaArchivos.Length; i++)
                        {
                            inicio = DateTime.Now;
                            file = new FileInfo(listaArchivos[i]);

                            ActualizaFechaProceso(GetProceso(file.Name), file.Length / 1024);

                            if (listaArchivos[i].Contains(ficheroGordo))
                                TrataFicheroGordo(listaArchivos[i], desdeLinea, desdeColumna, hastaColumna);

                            zip.ComprimirArchivo(listaArchivos[i], listaArchivos[i].Replace(".txt", ".zip"));
                            global.SaveProcess("ContratacionPortugal", "Comprimiendo " + file.Name, inicio, inicio, DateTime.Now);
                            file.Delete();

                        }
                    }
                    else
                    {
                        string from = p.GetValue("Buzon_envio_email");
                        string to = p.GetValue("supervisor_proceso");
                        string cc = null;
                        string subject = "Contratación Portugal (archivos normales) " + DateTime.Now.ToString("dd/MM/yyyy");
                        string body = "No se han encontrado ficheros en el extractor: "
                            + p.GetValue("Extractor_BTN_AAMMDD") + " " + ultimoDiaHabil;

                        //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                        EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
                        mes.SendMail(from, to, cc, subject, body);
                    }

                    listaArchivos = Directory.GetFiles(p.GetValue("Inbox"), "*.zip");

                    for (int i = 0; i < listaArchivos.Length; i++)
                    {
                        FileInfo fichero = new FileInfo(listaArchivos[i]);
                        ftp.Upload(p.GetValue("ruta_destino_FTP") + fichero.Name, listaArchivos[i]);
                        utilidades.Fichero.BorrarArchivo(listaArchivos[i]);
                    }


                }

                if (!ProcesoCompleto("BPO-Contratacion_SOLBTN_AAMMDD"))
                {
                    inicio = DateTime.Now;
                    ficheroLog.Add("Ejecutando extractor: " + p.GetValue("Extractor_BTN_AAMMDD") + " " + ultimoDiaHabil);
                    utilidades.Fichero.EjecutaComando(p.GetValue("Extractor_BTN_AAMMDD"), ultimoDiaHabil);
                    global.SaveProcess("ContratacionPortugal", "Extractor_BTN_AAMMDD", inicio, inicio, DateTime.Now);

                    string[] listaArchivos = Directory.GetFiles(p.GetValue("Inbox"));

                    ftp = new utilidades.UltimateFTP(
                      p.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                      p.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                      utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_pass", DateTime.Now, DateTime.Now), true),
                      p.GetValue("ftp_port", DateTime.Now, DateTime.Now));
                }


            }
            catch (Exception e)
            {
                ficheroLog.AddError("Portugal.EnviarArchivos: " + e.Message);
            }
        }

        public void Envia_SOL_BTN()
        {
            EndesaBusiness.utilidades.Global global = new EndesaBusiness.utilidades.Global();
            String ultimoDiaHabil = utilidades.Fichero.UltimoDiaHabil_YYMMDD();
            FileInfo file;
            DateTime inicio = new DateTime();
            string ficheroGordo = "";
            int desdeLinea = 0;
            int desdeColumna = 0;
            int hastaColumna = 0;

            utilidades.UltimateFTP ftp;
            utilidades.ZipUnZip zip;

            try
            {
                zip = new utilidades.ZipUnZip();
                ficheroGordo = p.GetValue("strSOLATRBTN_YYMMDD");

                inicio = DateTime.Now;
                ficheroLog.Add("Ejecutando extractor: " + p.GetValue("Extractor_BTN_AAMMDD") + " " + ultimoDiaHabil);
                utilidades.Fichero.EjecutaComando(p.GetValue("Extractor_BTN_AAMMDD"), ultimoDiaHabil);
                global.SaveProcess("ContratacionPortugal", "Extractor_BTN_AAMMDD", inicio, inicio, DateTime.Now);
                
                string[] listaArchivos = Directory.GetFiles(p.GetValue("Inbox"));


                ftp = new utilidades.UltimateFTP(
                       p.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                       p.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                       utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_pass", DateTime.Now, DateTime.Now), true),
                       p.GetValue("ftp_port", DateTime.Now, DateTime.Now));


                if (listaArchivos.Length == 2)
                {

                    for (int i = 0; i < listaArchivos.Length; i++)
                    {
                        inicio = DateTime.Now;
                        file = new FileInfo(listaArchivos[i]);

                        if (listaArchivos[i].Contains(ficheroGordo))
                            TrataFicheroGordo(listaArchivos[i], desdeLinea, desdeColumna, hastaColumna);

                        zip.ComprimirArchivo(listaArchivos[i], listaArchivos[i].Replace(".txt", ".zip"));
                        global.SaveProcess("ContratacionPortugal", "Comprimiendo " + file.Name, inicio, inicio, DateTime.Now);
                        file.Delete();

                    }
                }
                else
                {
                    string from = p.GetValue("Buzon_envio_email");
                    string to = p.GetValue("supervisor_proceso");
                    string cc = null;
                    string subject = "Contratación Portugal BTN (archivos grandes) " + DateTime.Now.ToString("dd/MM/yyyy");
                    string body = "No se han encontrado ficheros en el extractor: "
                        + p.GetValue("Extractor_BTN_AAMMDD") + " " + ultimoDiaHabil;

                    //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                    EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
                    mes.SendMail(from, to, cc, subject, body);
                }

                //inicio = DateTime.Now;
                //utilidades.Fichero.EjecutaComando(p.GetValue("SubirArchivosFTP"), ultimoDiaHabil);
                //global.SaveProcess("ContratacionPortugal", "SubirArchivosFTP_" + ultimoDiaHabil, inicio, inicio, DateTime.Now);

                //inicio = DateTime.Now;
                //utilidades.Fichero.EjecutaComando(p.GetValue("SubirArchivosFTP"), ultimoDiaHabil.Substring(2, 4));
                //global.SaveProcess("ContratacionPortugal", "SubirArchivosFTP_" + ultimoDiaHabil.Substring(2, 4), inicio, inicio, DateTime.Now);

                listaArchivos = Directory.GetFiles(p.GetValue("Inbox"), "*.zip");

                for (int i = 0; i < listaArchivos.Length; i++)
                {
                    FileInfo fichero = new FileInfo(listaArchivos[i]);
                    ftp.Upload(p.GetValue("ruta_destino_FTP") + fichero.Name, listaArchivos[i]);
                    utilidades.Fichero.BorrarArchivo(listaArchivos[i]);
                }

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Portugal.EnviarArchivos: " + e.Message);
            }
        }

        public void TrataFicheroGordo(string ficheroOrigen, Int64 desdeLinea, int desdeColumna, int hastaColumna)
        {
            long i = 0;
            string line;
            StringBuilder sb = new StringBuilder();
            string ficheroDestino = "";
            EndesaBusiness.utilidades.Global global = new EndesaBusiness.utilidades.Global();
            FileInfo file;
            DateTime inicio = new DateTime();
            utilidades.ZipUnZip zip;

            try
            {
                zip = new utilidades.ZipUnZip();

                inicio = DateTime.Now;

                ficheroDestino = ficheroOrigen.Replace(p.GetValue("strSOLATRBTN_YYMMDD"), p.GetValue("strSOLATRDPPARTIDO"));
                ficheroLog.Add("Partiendo fichero Origen: " + ficheroOrigen + " --> Fichero Destino: " + ficheroDestino);
                System.IO.StreamReader fileIn = new System.IO.StreamReader(ficheroOrigen);
                System.IO.StreamWriter fileOut = new System.IO.StreamWriter(ficheroDestino);

                Console.WriteLine("Reading " + ficheroOrigen + "...");
                while ((line = fileIn.ReadLine()) != null)
                {
                    i++;
                    if (i <= desdeLinea)
                    {
                        fileOut.WriteLine(line.Substring(desdeColumna - 1, hastaColumna - 1));
                    }
                }

                fileIn.Close();
                fileOut.Close();

                file = new FileInfo(ficheroDestino);
                ficheroLog.Add("Comprimiendo archivo " + ficheroDestino + " a " + ficheroDestino.Replace(".txt", ".zip"));

                zip.ComprimirArchivo(ficheroDestino, ficheroDestino.Replace(".txt", ".zip"));
                global.SaveProcess("ContratacionPortugal", "Comprimiendo " + file.Name, inicio, inicio, DateTime.Now);
                ficheroLog.Add("Borrando archivo " + file.FullName);
                file.Delete();

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Portugal.TrataFicheroGordo: " + e.Message);
            }
        }

        private void ActualizaFechaProceso(string proceso, long kb)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            DateTime d = new DateTime();

            try
            {
                d = DateTime.Now;
                
                strSql = "update cp_fechas_procesos set"
                    + " fecha = '" + d.ToString("yyyy-MM-dd") + "',"
                    + " kb = " + kb
                    + " where proceso = '" + proceso + "'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void ActualizaArchivoFechaProceso(string tabla, string archivo, DateTime fecha_ejecucion, long kb, string md5)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            DateTime d = new DateTime();

            try
            {
                d = DateTime.Now;

                strSql = "replace into " + tabla + " (archivo, fecha_ejecucion, kb, md5) values "
                    + "('" + archivo + "'," + "'" + fecha_ejecucion.ToString("yyyy-MM-dd") + "',"
                    + kb + ",'" + md5 + "')";                    
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private string GetProceso(string archivo)
        {
            foreach(KeyValuePair<string, EndesaEntity.contratacion.Tabla_cp_fechas_procesos> p in dic_fechas_procesos)
            {
                if(archivo.Contains(p.Key))
                    return p.Key;
            }
            return "";
        }

        private bool ProcesoCompleto(string extractor)
        {
            bool completo = true;
            List<DateTime> dateTimes = dic_fechas_procesos.Where(z => z.Value.extractor == extractor).Select(z => z.Value.fecha).ToList();
            foreach (DateTime date in dateTimes)
                if (date.Date < DateTime.Now.Date)
                    return false;

            return completo;

        }



    }
}
