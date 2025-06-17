using EndesaBusiness.servidores;
using K4os.Hash.xxHash;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.eexxi
{
    public class Aviso_a_COR : EndesaEntity.contratacion.xxi.XML_Datos
    {
        EndesaBusiness.utilidades.Param param;
        Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic;


        public Aviso_a_COR()
        {
            param = new EndesaBusiness.utilidades.Param("eexxi_param", MySQLDB.Esquemas.CON);
        }


        public void GuardaDatosXML_T102()
        {
            // Este proceso se hace como prueba para guardar los XML
            // t102 de los pasos ya rechazados

            Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic = Carga_Avisos_Rechazados();

            foreach(KeyValuePair<string, EndesaEntity.contratacion.xxi.XML_Datos> p in dic)
            {

            }


        }


        private Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> Carga_Avisos_Rechazados()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> d
                = new Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos>();

            try
            {
                strSql = "select codigodesolicitud, nif, razon_social, f_recepcion, f_prevista_alta,"
                    + " fichero from eexxi_aviso_paso_cor"
                    + " where rechazar = 'S'";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.XML_Datos c = new EndesaEntity.contratacion.xxi.XML_Datos();
                    c.codigoDeSolicitud = r["codigodesolicitud"].ToString();
                    c.fichero = r["fichero"].ToString();

                    EndesaEntity.contratacion.xxi.XML_Datos o;
                    if (!d.TryGetValue(c.codigoDeSolicitud, out o))
                        d.Add(c.codigoDeSolicitud, c);
                }
                db.CloseConnection();

                return d;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                 "Carga_Avisos",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
                return null;
            }
        }

        private Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> Carga_Avisos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> d
                = new Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos>();

            try
            {
                strSql = "select codigodesolicitud, nif, razon_social, f_recepcion, f_prevista_alta,"
                    + " fichero from eexxi_aviso_paso_cor"
                    + " where rechazar is null";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.XML_Datos c = new EndesaEntity.contratacion.xxi.XML_Datos();
                    c.codigoDeSolicitud = r["codigodesolicitud"].ToString();
                    c.fichero = r["fichero"].ToString();

                    EndesaEntity.contratacion.xxi.XML_Datos o;
                    if (!d.TryGetValue(c.codigoDeSolicitud, out o))                        
                        d.Add(c.codigoDeSolicitud, c);
                }
                db.CloseConnection();

                return d;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message,
                 "Carga_Avisos",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
                return null;
            }
        }

        private List<EndesaEntity.contratacion.xxi.Paso_a_COR> Carga_Avisos_Ventas(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            List<EndesaEntity.contratacion.xxi.Paso_a_COR> d
                = new List<EndesaEntity.contratacion.xxi.Paso_a_COR>();

            try
            {
                strSql = "select a.codigodesolicitud, s.cups, a.nif, a.razon_social, a.f_recepcion, a.f_prevista_alta,"
                    + " a.fichero from eexxi_aviso_paso_cor a inner join"
                    + " eexxi_solicitudes_t101 s on "
                    + " s.CodigoDeSolicitud = a.codigodesolicitud and"
                    + " s.Identificador = a.nif "
                    + " where rechazar is null and"
                    + " fecha_creacion >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_creacion <= '" + fh.AddDays(1).ToString("yyyy-MM-dd") + "'";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.Paso_a_COR c = new EndesaEntity.contratacion.xxi.Paso_a_COR();
                    c.codigo_solicitud = r["codigodesolicitud"].ToString();
                    c.cups = r["cups"].ToString();
                    c.nif = r["nif"].ToString();
                    c.razon_social = r["razon_social"].ToString();
                    c.f_recepcion = Convert.ToDateTime(r["f_recepcion"]);
                    c.f_prevista_alta = Convert.ToDateTime(r["f_prevista_alta"]);

                    d.Add(c);
                }
                db.CloseConnection();

                return d;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                 "Carga_Avisos",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
                return null;
            }
        }

        private List<EndesaEntity.contratacion.xxi.Paso_a_COR> Carga_Avisos(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            List<EndesaEntity.contratacion.xxi.Paso_a_COR> d
                = new List<EndesaEntity.contratacion.xxi.Paso_a_COR>();

            try
            {
                strSql = "select a.codigodesolicitud, s.cups, a.nif, a.razon_social, a.f_recepcion, a.f_prevista_alta,"
                    + " a.fichero from eexxi_aviso_paso_cor a inner join"
                    + " eexxi_solicitudes_t101 s on "
                    + " s.CodigoDeSolicitud = a.codigodesolicitud and"
                    + " s.Identificador = a.nif "
                    + " where rechazar is null and"
                    + " fecha_creacion >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_creacion <= '" + fh.AddDays(1).ToString("yyyy-MM-dd") + "'";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.Paso_a_COR c = new EndesaEntity.contratacion.xxi.Paso_a_COR();
                    c.codigo_solicitud = r["codigodesolicitud"].ToString();
                    c.cups = r["cups"].ToString();
                    c.nif = r["nif"].ToString();
                    c.razon_social = r["razon_social"].ToString();
                    c.f_recepcion = Convert.ToDateTime(r["f_recepcion"]);
                    c.f_prevista_alta = Convert.ToDateTime(r["f_prevista_alta"]);

                    d.Add(c);
                }
                db.CloseConnection();

                return d;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                 "Carga_Avisos",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
                return null;
            }
        }

        private string GetArchivo(string solicitud)
        {
            EndesaEntity.contratacion.xxi.XML_Datos o;
            if (dic.TryGetValue(solicitud, out o))
                return o.fichero;
            else
                return null;
        }

        private EndesaEntity.contratacion.xxi.XML_Datos GetXML(string solicitud)
        {

            EndesaEntity.contratacion.xxi.XML_Datos o;
            if (dic.TryGetValue(solicitud, out o))
            {
                EndesaEntity.contratacion.xxi.XML_Datos c = new EndesaEntity.contratacion.xxi.XML_Datos();
                c.codigoREEEmpresaEmisora = o.codigoREEEmpresaDestino;
                c.codigoREEEmpresaDestino = o.codigoREEEmpresaDestino;
                c.fichero = o.fichero; 
                c.secuencialDeSolicitud = o.secuencialDeSolicitud;

                return c;
            }
            else
                return null;
                
            
        }




        public void CargaExcelRespuestaCOR(string fichero)
        {

            
            int c = 1;
            int f = 1;
            FileStream fs;
            ExcelPackage excelPackage;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            List<EndesaEntity.contratacion.xxi.XML_Datos> lista
                = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
            Dictionary<string, EndesaEntity.contratacion.xxi.Cups_Solicitud> dic_solicitud_cups =
                new Dictionary<string, EndesaEntity.contratacion.xxi.Cups_Solicitud>();

            EndesaBusiness.contratacion.eexxi.EEXXI eEXXI = new EndesaBusiness.contratacion.eexxi.EEXXI();
            dic = Carga_Avisos();

            try
            {
                #region Cargamos el Excel de respuestas del área de Contratación

                fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();
                f = 1; // Porque la primera fila es la cabecera
                for (int i = 1; i < 5000; i++)
                {
                    c = 1;
                    f++;

                    if (workSheet.Cells[f, 1].Value == null)
                        break;

                    if (workSheet.Cells[f, 1].Value.ToString() == "")
                        break;

                    

                    EndesaEntity.contratacion.xxi.XML_Datos p = new EndesaEntity.contratacion.xxi.XML_Datos();
                    p.codigoDeSolicitud = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    p.cups = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    p.identificador = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    p.razonSocial = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    p.fechaSolicitud =  Convert.ToDateTime(workSheet.Cells[f, c].Value); c++;
                    p.fechaActivacion = Convert.ToDateTime(workSheet.Cells[f, c].Value); c++;
                    p.rechazar = Convert.ToString(workSheet.Cells[f, c].Value) == "RECHAZAR"; c++;                    

                    lista.Add(p);

                    EndesaEntity.contratacion.xxi.Cups_Solicitud k = new EndesaEntity.contratacion.xxi.Cups_Solicitud();
                    k.solicitud = p.codigoDeSolicitud;
                    k.cups = p.cups;


                    EndesaEntity.contratacion.xxi.Cups_Solicitud o;

                    if (!dic_solicitud_cups.TryGetValue(p.codigoDeSolicitud + "_" + p.cups, out o))
                    {
                        dic_solicitud_cups.Add(p.codigoDeSolicitud + "_" + p.cups, k);
                    }

                    
                }

                fs = null;
                excelPackage = null;

                #endregion

                #region Buscamos todos los t101 originales
                //Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic_t101 = 
                //    eEXXI.BuscaDatosSolicitudXML("eexxi_solicitudes_t101", "T1", "01", lista.Select(x => x.codigoDeSolicitud).ToList());

                Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic_t101 =
                    eEXXI.BuscaDatosSolicitudXML("eexxi_solicitudes_t101", "T1", "01", dic_solicitud_cups);

                foreach (EndesaEntity.contratacion.xxi.XML_Datos p in lista)
                {
                    EndesaEntity.contratacion.xxi.XML_Datos o;
                    if (dic_t101.TryGetValue(p.codigoDeSolicitud, out o))
                    {
                        p.codigoREEEmpresaEmisora = "0636";
                        p.codigoREEEmpresaDestino = o.codigoREEEmpresaEmisora;
                        p.secuencialDeSolicitud = o.secuencialDeSolicitud; 
                        p.fichero = o.fichero;
                        p.codigoDelProceso = "T1";
                        p.codigoDePaso = "02";

                    }
                }

                #endregion

                #region Generamos los XML del paso T102
                Update(lista);
                #endregion
               



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                 "CargaExcelRespuestaCOR",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
            }
        
        }

        public void GeneraExcelRespuestaCOR(string fichero, DateTime fd, DateTime fh)
        {


            int c = 1;
            int f = 1;
                    

            
            List<EndesaEntity.contratacion.xxi.Paso_a_COR> lista = Carga_Avisos(fd, fh);

            try
            {

                FileInfo file = new FileInfo(fichero);

                if (file.Exists)
                    file.Delete();

                FileInfo plantilla = new FileInfo(System.Environment.CurrentDirectory + param.GetValue("plantilla_excel_envio_a_cor"));

                ExcelPackage excelPackage;
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                excelPackage = new ExcelPackage(plantilla);
                var workSheet = excelPackage.Workbook.Worksheets["PASO_A_COR"];

                if (lista.Count > 0)
                {
                    foreach (EndesaEntity.contratacion.xxi.Paso_a_COR p in lista)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.codigo_solicitud; c++;
                        workSheet.Cells[f, c].Value = p.cups; c++;
                        workSheet.Cells[f, c].Value = p.nif; c++;
                        workSheet.Cells[f, c].Value = p.razon_social; c++;
                        workSheet.Cells[f, c].Value = p.f_recepcion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                        workSheet.Cells[f, c].Value = p.f_prevista_alta;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                    }

                    var allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();
                   

                }

                excelPackage.SaveAs(file);
                excelPackage = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                 "CargaExcelRespuestaCOR",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
            }

        }
        public void GeneraExcelAvisoVentas(string fichero, DateTime fd, DateTime fh)
        {
            int c = 1;
            int f = 1;

            List<EndesaEntity.contratacion.xxi.Paso_a_COR> lista = Carga_Avisos_Ventas(fd, fh);

            try
            {

                FileInfo file = new FileInfo(fichero);

                if (file.Exists)
                    file.Delete();

                FileInfo plantilla = new FileInfo(System.Environment.CurrentDirectory + param.GetValue("plantilla_excel_envio_a_cor"));

                ExcelPackage excelPackage;
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                excelPackage = new ExcelPackage(plantilla);
                var workSheet = excelPackage.Workbook.Worksheets["PASO_A_COR"];

                if (lista.Count > 0)
                {
                    foreach (EndesaEntity.contratacion.xxi.Paso_a_COR p in lista)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.codigo_solicitud; c++;
                        workSheet.Cells[f, c].Value = p.cups; c++;
                        workSheet.Cells[f, c].Value = p.nif; c++;
                        workSheet.Cells[f, c].Value = p.razon_social; c++;
                        workSheet.Cells[f, c].Value = p.f_recepcion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                        workSheet.Cells[f, c].Value = p.f_prevista_alta;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                    }

                    var allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();


                }

                excelPackage.SaveAs(file);
                excelPackage = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                 "CargaExcelRespuestaCOR",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
            }

        }

        private void Update(List<EndesaEntity.contratacion.xxi.XML_Datos> lista)
        {
            DirectoryInfo dirSalida;
            int secuencial;
            string fileName = "";
            string fechaHora = "";

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            List<string> lista_archivos_t102 =
                new List<string>();

            List<EndesaEntity.contratacion.xxi.XML_Datos> lista_t102 =
                new List<EndesaEntity.contratacion.xxi.XML_Datos>();

            EEXXI xxi = new EEXXI();

            //EndesaBusiness.utilidades.ZIP zip = new EndesaBusiness.utilidades.ZIP();
            utilidades.ZipUnZip zip = new utilidades.ZipUnZip();

            StringBuilder cuerpo = new StringBuilder();
            EndesaBusiness.office.MailCompose mail = new EndesaBusiness.office.MailCompose();

            try
            {
                dirSalida = new DirectoryInfo(param.GetValue("RutaSalidaXML_T102"));
                if (!dirSalida.Exists)
                    dirSalida.Create();

                BorrarContenidoDirectorio(param.GetValue("RutaSalidaXML_T102"));

                for (int i = 0; i < lista.Count; i++)
                {

                    if (lista[i].rechazar)
                    {
                        // lista_archivos_t102.Add(lista[i].fichero);
                        lista[i].fichero = CreaXML_Rechazo(lista[i]);
                        lista_t102.Add(lista[i]);
                    }
                                      

                    strSql = "update eexxi_aviso_paso_cor set"
                                    + " last_update_by = '" + System.Environment.UserName.ToUpper() + "'";

                    strSql += ", rechazar = '" + (lista[i].rechazar ? "S" : "N") + "'";
                    strSql += ", fecha_recepcion = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                    strSql += ", fecha_generacion_xml = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'";

                    strSql += " where codigodesolicitud = '" + lista[i].codigoDeSolicitud + "'"
                    + " and nif = '" + lista[i].identificador + "'";

                    db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                }

                #region Guardamos los XML de t102 en BBDD
                xxi.GuardadoBBDD(lista_t102, "eexxi_solicitudes");
                #endregion


                // Comprimimos los XML y los enviamos por mail
                fechaHora = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                fileName = param.GetValue("RutaSalidaXML_T102") 
                        + "T1_02_"                        
                        + fechaHora
                        + ".zip";
                FileInfo archivo = new FileInfo(fileName);

                //zip.ComprimirVarios(param.GetValue("RutaSalidaXML_T102"),
                //    ".*\\.(xml)$", archivo.FullName);

                zip.ComprimirVarios(param.GetValue("RutaSalidaXML_T102") + "*.xml" ,archivo.FullName);


                if (param.GetValue("Aviso_paso_a_COR") == "S")
                {

                    office.SendMail mail_AAPP = new office.SendMail(param.GetValue("buzon_mail", DateTime.Now, DateTime.Now));

                    cuerpo.Append(System.Environment.NewLine);
                    cuerpo.Append(DateTime.Now.Hour > 14 ? "Buenas tardes:" : "Buenos días:");
                    cuerpo.Append(System.Environment.NewLine);
                    cuerpo.Append("Solicitamos subida de los ficheros adjuntos a IMEG.");
                    cuerpo.Append(System.Environment.NewLine);
                    cuerpo.Append("Por favor, comuniquen la correcta subida del mismo.");
                    cuerpo.Append(System.Environment.NewLine);
                    cuerpo.Append("Saludos.");

                    mail_AAPP.para.Add(param.GetValue("Mail_XML_Rechazos"));
                    mail_AAPP.asunto = param.GetValue("subida_xml");
                    mail_AAPP.htmlCuerpo = cuerpo.ToString();
                    mail_AAPP.adjuntos.Add(archivo.FullName);

                    if (param.GetValue("enviar_mail_Paso_a_COR", DateTime.Now, DateTime.Now) != "S")
                        mail_AAPP.Save();
                    else
                        mail_AAPP.Send();

                    archivo.Delete();
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                "Update",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }


        }

        private string CreaXML_Rechazo(EndesaEntity.contratacion.xxi.XML_Datos xml)
        {
            
            EndesaBusiness.xml.XMLFunciones xml_t102 = new EndesaBusiness.xml.XMLFunciones();
            EndesaEntity.contratacion.xxi.XML_Datos t101;
            int secuencial;

            try
            {
                secuencial = Convert.ToInt32(param.GetValue("secuencial_solicitud")) + 1;               

                FileInfo ficheroSalida = new FileInfo(param.GetValue("RutaSalidaXML_T102")
                    + xml.codigoREEEmpresaEmisora + "_"
                    + xml.codigoREEEmpresaDestino + "_"
                    + "T1_02_"
                    + xml.cups + "_"
                    + "01_"
                    + secuencial.ToString().PadLeft(4, '0')
                    + ".xml");
                    
                xml_t102.CreaXML_T102(ficheroSalida,xml);
                Thread.Sleep(1000);
                GuardaNumSecuencialTemporal(secuencial);
                return ficheroSalida.Name;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                "CreaXML_Rechazo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                return null;
            }
            

            

        }

        public void GuardadoBBDD_Paso_a_COR(List<EndesaEntity.contratacion.xxi.XML_Datos> lista)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            bool firstOnly = true;
            int num_reg = 0;

            foreach (EndesaEntity.contratacion.xxi.XML_Datos xml in lista)
            {
                if (firstOnly)
                {
                    strSql = "replace into eexxi_aviso_paso_cor (codigodesolicitud, nif," +
                        " razon_social, f_recepcion, f_prevista_alta, rechazar," +
                        " fecha_creacion, fecha_recepcion, fecha_generacion_xml," +
                        " fichero, created_by) values ";
                    firstOnly = false;
                }

                num_reg++;

                #region Campos                

                if (xml.codigoDeSolicitud != null)
                    strSql += "('" + xml.codigoDeSolicitud + "'";
                else
                    strSql += "(null";

                if (xml.identificador != null)
                    strSql += ", '" + xml.identificador + "'";
                else
                    strSql += ", null";

                if (xml.razonSocial != null)
                    strSql += ", '" + xml.razonSocial + "'";
                else
                    strSql += ", null";

                if (xml.fechaSolicitud != null)
                    strSql += ", '" + xml.fechaSolicitud.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ", null";

                if (xml.fechaActivacion != null)
                    strSql += ", '" + xml.fechaActivacion.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ", null";

                // rechazar
                strSql += ", null";
                // fecha_creacion
                strSql += ", '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                // fecha_recepcion
                strSql += ", null";
                // fecha_generacion_xml
                strSql += ", null";
                strSql += ", '" + xml.fichero + "'" + ", '" + System.Environment.UserName + "'),";
                #endregion

                if (num_reg > 250)
                {
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    num_reg = 0;
                    strSql = "";
                    firstOnly = true;
                }

            }

            if (num_reg > 0)
            {
                db = new EndesaBusiness.servidores.MySQLDB(EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                num_reg = 0;
                strSql = "";
            }


        }
        public void GuardadoBBDD_Aviso_Ventas(List<EndesaEntity.contratacion.xxi.XML_Datos> lista)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            bool firstOnly = true;
            int num_reg = 0;
            string direccion_ps = "";
            string direccion_cliente = "";
            Dictionary<string, string> dic_nifs = new Dictionary<string, string>();
            string rt = "";

            EndesaBusiness.global.Provincias provincias = 
                new EndesaBusiness.global.Provincias("eexxi_param_provincias", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);


            // Obtenemos los nif para cartera

            foreach (EndesaEntity.contratacion.xxi.XML_Datos xml in lista)
            {
                string o;
                if(!dic_nifs.TryGetValue(xml.identificador, out o))
                    dic_nifs.Add(xml.identificador, xml.identificador);                
            }

            EndesaBusiness.cartera.Cartera_SalesForce sf =
                new cartera.Cartera_SalesForce(dic_nifs.Values.ToList());


            foreach (EndesaEntity.contratacion.xxi.XML_Datos xml in lista)
            {
                if (firstOnly)
                {
                    strSql = "replace into eexxi_aviso_ventas (codigodesolicitud, nif, cups," +
                        " razon_social, direccion_ps, direccion_cliente, f_recepcion, f_prevista_alta," +                        
                        " fichero, rt, created_by, created_date) values ";
                    firstOnly = false;
                }

                num_reg++;

                #region Campos                

                if (xml.codigoDeSolicitud != null)
                    strSql += "('" + xml.codigoDeSolicitud + "'";
                else
                    strSql += "(null";

                if (xml.identificador != null)
                    strSql += ", '" + xml.identificador + "'";
                else
                    strSql += ", null";

                if (xml.cups != null)
                    strSql += ", '" + xml.cups + "'";
                else
                    strSql += ", null";

                if (xml.razonSocial != null)
                    strSql += ", '" + xml.razonSocial + "'";
                else
                    strSql += ", null";


                // Direcciones

                direccion_ps = xml.linea1DeLaDireccionExterna
                    + " " + xml.linea2DeLaDireccionExterna;

                strSql += ", '" + direccion_ps + "'";

                direccion_cliente = xml.calleCliente + ", "
                    + xml.numeroFincaCliente
                    + " " + xml.codPostalCliente
                    + " " + provincias.DesProvincia(xml.codPostalCliente);

                strSql += ", '" + direccion_cliente + "'";


                if (xml.fechaSolicitud != null)
                    strSql += ", '" + xml.fechaSolicitud.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ", null";

                if (xml.fechaActivacion != null)
                    strSql += ", '" + xml.fechaActivacion.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ", null";

                strSql += ", '" + xml.fichero + "'";


                if (sf.ExisteCartera(xml.identificador))
                {
                    rt = sf.email_responsable_rt;
                    strSql += ", '" + rt + "'";
                }else                    
                    strSql += ", null";

                // created_by
                strSql += ", '" + System.Environment.UserName + "'";
                // created_date
                strSql += ", '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'),";
                 
                #endregion

                if (num_reg > 250)
                {
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    num_reg = 0;
                    strSql = "";
                    firstOnly = true;
                }

            }

            if (num_reg > 0)
            {
                db = new EndesaBusiness.servidores.MySQLDB(EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                num_reg = 0;
                strSql = "";
            }


        }

        public void GeneraMails_Paso_a_COR(List<EndesaEntity.contratacion.xxi.XML_Datos> lista)
        {
            // Este archivo Excel se le pasará a Cobros
            // En función de la aceptación por cobros o no
            // Posteriormente se generará Rechazo.

            StringBuilder cuerpo = new StringBuilder();
            EndesaBusiness.office.MailCompose mail = new EndesaBusiness.office.MailCompose();

            DirectoryInfo dirSalida;
            //FileInfo nombreSalidaExcel_AAPP;
            //FileInfo nombreSalidaExcel_EEPP;
            FileInfo nombreSalidaExcel;
            int f = 1;
            int c = 1;
            

            try
            {
                // Ruta de la plantilla 
                FileInfo plantilla = new FileInfo(System.Environment.CurrentDirectory +  param.GetValue("plantilla_excel_envio_a_cor"));

                dirSalida = new DirectoryInfo(param.GetValue("RutaSalidaExcelAltas"));

                if (!dirSalida.Exists)
                    dirSalida.Create();

                // Salida Excel para Administraciones Publicas
                //nombreSalidaExcel_AAPP = new FileInfo(dirSalida.FullName 
                //    + param.GetValue("Excel_COR_AAPP") + "_" 
                //    + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

                // Salida Excel para Empresas Privadas
                //nombreSalidaExcel_EEPP = new FileInfo(dirSalida.FullName
                //    + param.GetValue("Excel_COR_EEPP") + "_"
                //    + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

                // Salida Excel conjunto
                nombreSalidaExcel = new FileInfo(dirSalida.FullName
                    + param.GetValue("Excel_COR") + "_"
                    + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");


                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(plantilla);
                var workSheet = excelPackage.Workbook.Worksheets["PASO_A_COR"];

                var headerCells = workSheet.Cells[1, 1, 1, 25];
                var headerFont = headerCells.Style.Font;
                                             

                List<EndesaEntity.contratacion.xxi.XML_Datos> lista_aapp = lista;

                if (lista_aapp.Count > 0)
                {
                    f = 1;
                    for (int i = 0; i < lista_aapp.Count; i++)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = lista_aapp[i].codigoDeSolicitud; c++;
                        workSheet.Cells[f, c].Value = lista_aapp[i].cups; c++;
                        workSheet.Cells[f, c].Value = lista_aapp[i].identificador; c++;
                        workSheet.Cells[f, c].Value = lista_aapp[i].razonSocial; c++;
                        workSheet.Cells[f, c].Value = lista_aapp[i].fechaSolicitud;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                        workSheet.Cells[f, c].Value = lista_aapp[i].fechaActivacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                    }

                    var allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();
                    //excelPackage.SaveAs(nombreSalidaExcel_AAPP);
                    excelPackage.SaveAs(nombreSalidaExcel);
                    excelPackage = null;

                    if (param.GetValue("Aviso_paso_a_COR") == "S")
                    {
                                                
                        office.SendMail mail_AAPP = new office.SendMail(param.GetValue("buzon_mail", DateTime.Now, DateTime.Now));
                        cuerpo = null;
                        cuerpo = new StringBuilder();

                        cuerpo.Append(System.Environment.NewLine);
                        cuerpo.Append(DateTime.Now.Hour > 14 ? "Buenas tardes:" : "Buenos días:");
                        cuerpo.Append(System.Environment.NewLine);
                        cuerpo.Append("Adjunto enviamos posibles pasos a COR que nos informan las ");                        
                        cuerpo.Append("distribuidoras, por favor solicitamos que nos remitan en el día de hoy si ");
                        cuerpo.Append("debemos enviar algún fichero de rechazo.");
                        cuerpo.Append(System.Environment.NewLine);
                        cuerpo.Append("Saludos.");

                        mail_AAPP.para.Add(param.GetValue("Mail_Paso_a_COR_AAPP"));
                        mail_AAPP.cc.Add(param.GetValue("Mail_Paso_a_COR_AAPP_CC"));
                        //mail_AAPP.asunto = "Paso a COR Administración Pública";
                        mail_AAPP.asunto = "Paso a COR";
                        mail_AAPP.htmlCuerpo = cuerpo.ToString();
                        //mail_AAPP.adjuntos.Add(nombreSalidaExcel_AAPP.FullName);
                        mail_AAPP.adjuntos.Add(nombreSalidaExcel.FullName);

                        if (param.GetValue("enviar_mail_Paso_a_COR", DateTime.Now, DateTime.Now) != "S")
                            mail_AAPP.Save();
                        else
                            mail_AAPP.Send();

                        //nombreSalidaExcel_AAPP.Delete();
                        nombreSalidaExcel.Delete();
                    }
                }

                


                //f = 1;
                //c = 1;

                //ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                //excelPackage = new ExcelPackage(plantilla);
                //workSheet = excelPackage.Workbook.Worksheets["PASO_A_COR"];

                //headerCells = workSheet.Cells[1, 1, 1, 25];
                //headerFont = headerCells.Style.Font;               


                //List<EndesaEntity.contratacion.xxi.XML_Datos> lista_eepp
                //    = lista.Where(z => (z.identificador.Substring(0, 1) != "P")
                //    && (z.identificador.Substring(0, 1) != "Q")
                //    && (z.identificador.Substring(0, 1) != "S")).ToList();

                //if(lista_eepp.Count > 0)
                //{
                //    for (int i = 0; i < lista_eepp.Count; i++)
                //    {
                //        c = 1;
                //        f++;
                //        workSheet.Cells[f, c].Value = lista_eepp[i].codigoDeSolicitud; c++;
                //        workSheet.Cells[f, c].Value = lista_eepp[i].cups; c++;
                //        workSheet.Cells[f, c].Value = lista_eepp[i].identificador; c++;
                //        workSheet.Cells[f, c].Value = lista_eepp[i].razonSocial; c++;
                //        workSheet.Cells[f, c].Value = lista_eepp[i].fechaSolicitud;
                //        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                //        workSheet.Cells[f, c].Value = lista_eepp[i].fechaActivacion;
                //        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                //    }

                //    var allCells2 = workSheet.Cells[1, 1, f, c];
                //    allCells2.AutoFitColumns();
                //    excelPackage.SaveAs(nombreSalidaExcel_EEPP);
                //    excelPackage = null;

                //    if (param.GetValue("Aviso_paso_a_COR") == "S")
                //    {
                //        office.SendMail mail_EEPP = new office.SendMail(param.GetValue("buzon_mail", DateTime.Now, DateTime.Now));

                //        cuerpo = null;
                //        cuerpo = new StringBuilder();

                //        cuerpo.Append(System.Environment.NewLine);
                //        cuerpo.Append(DateTime.Now.Hour > 14 ? "Buenas tardes:" : "Buenos días:");
                //        cuerpo.Append(System.Environment.NewLine);
                //        cuerpo.Append("Adjunto enviamos posibles pasos a COR que nos informan las ");
                //        cuerpo.Append("distribuidoras, por favor solicitamos que nos remitan en el día de hoy si ");
                //        cuerpo.Append("debemos enviar algún fichero de rechazo.");
                //        cuerpo.Append(System.Environment.NewLine);
                //        cuerpo.Append("Saludos.");

                //        mail_EEPP.para.Add(param.GetValue("Mail_Paso_a_COR_EEPP"));
                //        mail_EEPP.asunto = "Paso a COR Empresa Privada";
                //        mail_EEPP.htmlCuerpo = cuerpo.ToString();
                //        mail_EEPP.adjuntos.Add(nombreSalidaExcel_EEPP.FullName);

                //        if (param.GetValue("enviar_mail_Paso_a_COR", DateTime.Now, DateTime.Now) != "S")
                //            mail_EEPP.Save();
                //        else
                //            mail_EEPP.Send();

                //        nombreSalidaExcel_EEPP.Delete();
                //    }
                //}

                

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message,
                   "GeneraMails_Paso_a_COR",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }

           

        }


        public void GeneraMails_Aviso_Ventas()
        {
            // Este archivo Excel se le pasará a Ventas           

            StringBuilder cuerpo = new StringBuilder();
            EndesaBusiness.office.MailCompose mail = new EndesaBusiness.office.MailCompose();

            DirectoryInfo dirSalida;
            //FileInfo nombreSalidaExcel_AAPP;
            //FileInfo nombreSalidaExcel_EEPP;
            FileInfo nombreSalidaExcel;
            int f = 1;
            int c = 1;

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            List<string> lista_rt = new List<string>();


            try
            {
                // Ruta de la plantilla 
                FileInfo plantilla = new FileInfo(System.Environment.CurrentDirectory + param.GetValue("plantilla_excel_envio_a_ventas"));

                dirSalida = new DirectoryInfo(param.GetValue("RutaSalidaExcelAltas"));

                if (!dirSalida.Exists)
                    dirSalida.Create();


                // Obtemos la lista de los RT

                strSql = "select rt"
                    + " from eexxi_aviso_ventas where enviado = 'N'"
                    + " group by rt";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if(r["rt"] != System.DBNull.Value)
                        lista_rt.Add(r["rt"].ToString());                   
                        
                }
                db.CloseConnection();
                if (lista_rt.Count > 0)
                {
                    foreach (string rt in lista_rt)
                    {

                        nombreSalidaExcel = new FileInfo(dirSalida.FullName
                        + param.GetValue("Excel_Ventas") + "_"
                        + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                        ExcelPackage excelPackage = new ExcelPackage(plantilla);
                        var workSheet = excelPackage.Workbook.Worksheets["Aviso Altas"];

                        var headerCells = workSheet.Cells[1, 1, 1, 25];
                        var headerFont = headerCells.Style.Font;

                        strSql = "select codigodesolicitud, cups, nif, razon_social, direccion_ps,"
                        + " direccion_cliente, f_recepcion, f_prevista_alta, rt"
                        + " from eexxi_aviso_ventas where enviado = 'N' and "
                        + " rt = '" + rt + "'";
                        db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(strSql, db.con);
                        r = command.ExecuteReader();

                        f = 1;
                        while (r.Read())
                        {
                            c = 1;
                            f++;
                            if (r["codigodesolicitud"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["codigodesolicitud"].ToString();
                            c++;

                            if (r["cups"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["cups"].ToString();
                            c++;

                            if (r["direccion_ps"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["direccion_ps"].ToString();
                            c++;

                            if (r["nif"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["nif"].ToString();
                            c++;

                            if (r["razon_social"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["razon_social"].ToString();
                            c++;

                            if (r["direccion_cliente"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["direccion_cliente"].ToString();
                            c++;

                            if (r["f_recepcion"] != System.DBNull.Value)
                            {
                                workSheet.Cells[f, c].Value = Convert.ToDateTime(r["f_recepcion"]);
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;

                            if (r["f_prevista_alta"] != System.DBNull.Value)
                            {
                                workSheet.Cells[f, c].Value = Convert.ToDateTime(r["f_prevista_alta"]);
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;

                        }
                        db.CloseConnection();

                        var allCells = workSheet.Cells[1, 1, f, c];
                        allCells.AutoFitColumns();
                        //excelPackage.SaveAs(nombreSalidaExcel_AAPP);
                        excelPackage.SaveAs(nombreSalidaExcel);
                        excelPackage = null;

                        if (param.GetValue("Aviso_a_Ventas") == "S")
                        {
                            cuerpo = null;
                            cuerpo = new StringBuilder();

                            office.SendMail mail_AAPP = new office.SendMail(param.GetValue("buzon_mail", DateTime.Now, DateTime.Now));

                            cuerpo.Append(System.Environment.NewLine);
                            cuerpo.Append(DateTime.Now.Hour > 14 ? "Buenas tardes:" : "Buenos días:");
                            cuerpo.Append(System.Environment.NewLine);
                            cuerpo.Append("Adjunto enviamos posibles altas que nos informan las ");
                            cuerpo.Append("distribuidoras. ");
                            cuerpo.Append(System.Environment.NewLine);
                            cuerpo.Append("Saludos.");

                            mail_AAPP.para.Add(rt);
                            //mail_AAPP.cc.Add(param.GetValue("Mail_Paso_a_COR_AAPP_CC"));
                            //mail_AAPP.asunto = "Paso a COR Administración Pública";
                            mail_AAPP.asunto = "Aviso Altas";
                            mail_AAPP.htmlCuerpo = cuerpo.ToString();
                            //mail_AAPP.adjuntos.Add(nombreSalidaExcel_AAPP.FullName);
                            mail_AAPP.adjuntos.Add(nombreSalidaExcel.FullName);

                            if (param.GetValue("enviar_mail_Aviso_Altas", DateTime.Now, DateTime.Now) != "S")
                                mail_AAPP.Save();
                            else
                                mail_AAPP.Send();

                            //nombreSalidaExcel_AAPP.Delete();
                            nombreSalidaExcel.Delete();

                            // Marcamos los registros como enviados

                            strSql = "update eexxi_aviso_ventas set enviado = 'S' where"
                                + " enviado = 'N' and"
                                + " rt = '" + rt + "'";
                            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                            command = new MySqlCommand(strSql, db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();

                        }

                    }

                    MessageBox.Show("Se han generado "
                        + lista_rt.Count.ToString("N0")
                        + " mails.",
                      "Mails aviso a Ventas",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information);

                }
                else
                {
                    MessageBox.Show("No hay RT para generar mails",
                       "Mails aviso a Ventas",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Information);
                }




            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                   "GeneraMails_Paso_a_COR",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }



        }

        public DateTime UltimoDiaCarga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime max_fecha = new DateTime();
            try
            {
                strSql = "SELECT MAX(created_date) max_fecha FROM eexxi_aviso_ventas";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    max_fecha = Convert.ToDateTime(r["max_fecha"]);
                }
                db.CloseConnection();
                return max_fecha.Date;
            }
            catch(Exception ex)
            {
                return DateTime.MinValue.Date;
            }

        }

        public void GeneraMails_Aviso_Ventas(DateTime fd, DateTime fh)
        {
            // Este archivo Excel se le pasará a Ventas           

            StringBuilder cuerpo = new StringBuilder();
            EndesaBusiness.office.MailCompose mail = new EndesaBusiness.office.MailCompose();

            DirectoryInfo dirSalida;
            //FileInfo nombreSalidaExcel_AAPP;
            //FileInfo nombreSalidaExcel_EEPP;
            FileInfo nombreSalidaExcel;
            int f = 1;
            int c = 1;

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            List<string> lista_rt = new List<string>();


            try
            {
                // Ruta de la plantilla 
                FileInfo plantilla = new FileInfo(System.Environment.CurrentDirectory + param.GetValue("plantilla_excel_envio_a_ventas"));

                dirSalida = new DirectoryInfo(param.GetValue("RutaSalidaExcelAltas"));

                if (!dirSalida.Exists)
                    dirSalida.Create();


                // Obtemos la lista de los RT

                strSql = "select rt"
                    + " from eexxi_aviso_ventas where"
                    + " (created_date >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " created_date <= '" + fh.AddDays(1).ToString("yyyy-MM-dd") + "')"
                    + " group by rt";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["rt"] != System.DBNull.Value)
                        lista_rt.Add(r["rt"].ToString());

                }
                db.CloseConnection();
                if(lista_rt.Count > 0)
                {
                    foreach (string rt in lista_rt)
                    {

                        nombreSalidaExcel = new FileInfo(dirSalida.FullName
                        + param.GetValue("Excel_Ventas") + "_"
                        + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                        ExcelPackage excelPackage = new ExcelPackage(plantilla);
                        var workSheet = excelPackage.Workbook.Worksheets["Aviso Altas"];

                        var headerCells = workSheet.Cells[1, 1, 1, 25];
                        var headerFont = headerCells.Style.Font;

                        strSql = "select codigodesolicitud, cups, nif, razon_social, direccion_ps,"
                        + " direccion_cliente, f_recepcion, f_prevista_alta, rt"
                        + " from eexxi_aviso_ventas where "
                         + " (created_date >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                        + " created_date <= '" + fh.AddDays(1).ToString("yyyy-MM-dd") + "') and"
                        + " rt = '" + rt + "'";
                        db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(strSql, db.con);
                        r = command.ExecuteReader();

                        f = 1;
                        while (r.Read())
                        {
                            c = 1;
                            f++;
                            if (r["codigodesolicitud"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["codigodesolicitud"].ToString();
                            c++;

                            if (r["cups"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["cups"].ToString();
                            c++;

                            if (r["direccion_ps"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["direccion_ps"].ToString();
                            c++;

                            if (r["nif"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["nif"].ToString();
                            c++;

                            if (r["razon_social"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["razon_social"].ToString();
                            c++;

                            if (r["direccion_cliente"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["direccion_cliente"].ToString();
                            c++;

                            if (r["f_recepcion"] != System.DBNull.Value)
                            {
                                workSheet.Cells[f, c].Value = Convert.ToDateTime(r["f_recepcion"]);
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;

                            if (r["f_prevista_alta"] != System.DBNull.Value)
                            {
                                workSheet.Cells[f, c].Value = Convert.ToDateTime(r["f_prevista_alta"]);
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;

                        }
                        db.CloseConnection();

                        var allCells = workSheet.Cells[1, 1, f, c];
                        allCells.AutoFitColumns();
                        //excelPackage.SaveAs(nombreSalidaExcel_AAPP);
                        excelPackage.SaveAs(nombreSalidaExcel);
                        excelPackage = null;

                        if (param.GetValue("Aviso_a_Ventas") == "S")
                        {
                            cuerpo = null;
                            cuerpo = new StringBuilder();

                            office.SendMail mail_AAPP = new office.SendMail(param.GetValue("buzon_mail", DateTime.Now, DateTime.Now));

                            cuerpo.Append(System.Environment.NewLine);
                            cuerpo.Append(DateTime.Now.Hour > 14 ? "Buenas tardes:" : "Buenos días:");
                            cuerpo.Append(System.Environment.NewLine);
                            cuerpo.Append("Adjunto enviamos posibles altas que nos informan las ");
                            cuerpo.Append("distribuidoras. ");
                            cuerpo.Append(System.Environment.NewLine);
                            cuerpo.Append("Saludos.");

                            mail_AAPP.para.Add(rt);
                            //mail_AAPP.cc.Add(param.GetValue("Mail_Paso_a_COR_AAPP_CC"));
                            //mail_AAPP.asunto = "Paso a COR Administración Pública";
                            mail_AAPP.asunto = "Aviso Altas";
                            mail_AAPP.htmlCuerpo = cuerpo.ToString();
                            //mail_AAPP.adjuntos.Add(nombreSalidaExcel_AAPP.FullName);
                            mail_AAPP.adjuntos.Add(nombreSalidaExcel.FullName);

                            if (param.GetValue("enviar_mail_Aviso_Altas", DateTime.Now, DateTime.Now) != "S")
                                mail_AAPP.Save();
                            else
                                mail_AAPP.Send();

                            //nombreSalidaExcel_AAPP.Delete();
                            nombreSalidaExcel.Delete();

                            // Marcamos los registros como enviados

                            strSql = "update eexxi_aviso_ventas set enviado = 'S' where"
                                + " enviado = 'N' and"
                                + " rt = '" + rt + "'";
                            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                            command = new MySqlCommand(strSql, db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();

                        }

                    }

                    MessageBox.Show("Se han generado " 
                        + lista_rt.Count.ToString("N0")
                        + " mails.",
                      "Mails aviso a Ventas",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information);

                }
                else
                {
                    MessageBox.Show("No hay RT para generar mails",
                       "Mails aviso a Ventas",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Information);
                }
            

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                   "GeneraMails_Paso_a_COR",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }



        }

        private void GuardaNumSecuencialTemporal(int secuencial_solicitud)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            strSql = "update eexxi_param set value = '" + secuencial_solicitud + "'"
                + " where code = 'secuencial_solicitud'";
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            // Volvemos a cargar parametros para tener el ultimo valor en memoria
            param = new utilidades.Param("eexxi_param", servidores.MySQLDB.Esquemas.CON);

        }

        private void BorrarContenidoDirectorio(string directorio)
        {
            string[] listaArchivos;
            FileInfo file;

            listaArchivos = Directory.GetFiles(directorio);
            for (int i = 0; i < listaArchivos.Count(); i++)
            {
                file = new FileInfo(listaArchivos[i]);
                file.Delete();
            }
        }
    }
}
