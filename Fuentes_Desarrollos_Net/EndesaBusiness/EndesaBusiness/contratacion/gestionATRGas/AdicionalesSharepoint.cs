using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion.gestionATRGas
{
    public class AdicionalesSharepoint
    {
        sharePoint.Utilidades sharePoint;
        logs.Log ficheroLog;
        utilidades.Param param;
        bool hay_error = true;
        Dictionary<int, EndesaEntity.contratacion.gas.AdicionalSharepoint> dic_adicionales_procesados;
        public AdicionalesSharepoint()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_AdicionalesSharepoint");
            param = new utilidades.Param("atrgas_param", servidores.MySQLDB.Esquemas.CON);
            sharePoint = new sharePoint.Utilidades();
            
        }

        public void Proceso()
        {
            BorrarContenidoDirectorio(param.GetValue("AZUCARERA_Carpeta_Descarga", DateTime.Now, DateTime.Now));

            if (!hay_error)
            {
                DescargaArchivoExcelSharepoint();
                dic_adicionales_procesados = CargaAdicionalesProcesados();                
            }

            if (!hay_error)
            {                
                GeneraXML(LeeExcel());
            }


        }

        private Dictionary<int, EndesaEntity.contratacion.gas.AdicionalSharepoint> CargaAdicionalesProcesados()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            Dictionary<int, EndesaEntity.contratacion.gas.AdicionalSharepoint> d =
                new Dictionary<int, EndesaEntity.contratacion.gas.AdicionalSharepoint>();

            try
            {

                ficheroLog.Add("Cargando datos de la tabla atrgas_solicitudes_sharepoint");
                Console.WriteLine("Cargando datos de la tabla atrgas_solicitudes_sharepoint");


                strSql = "select id, hora_de_inicio, hora_de_finalizacion, correo_electronico,"
                    + " nombre, cliente, cups, producto, fecha_inicio"
                    + " from atrgas_solicitudes_sharepoint";                
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.gas.AdicionalSharepoint c = new EndesaEntity.contratacion.gas.AdicionalSharepoint();
                    if (r["id"] != System.DBNull.Value)
                        c.id = Convert.ToInt32(r["id"]);

                    if (r["hora_de_inicio"] != System.DBNull.Value)
                        c.hora_de_inicio = Convert.ToDateTime(r["hora_de_inicio"]);

                    if (r["hora_de_finalizacion"] != System.DBNull.Value)
                        c.hora_de_finalizacion = Convert.ToDateTime(r["hora_de_finalizacion"]);

                    if (r["correo_electronico"] != System.DBNull.Value)
                        c.correo_electronico = Convert.ToString(r["correo_electronico"]);

                    if (r["nombre"] != System.DBNull.Value)
                        c.nombre = Convert.ToString(r["nombre"]);

                    if (r["cliente"] != System.DBNull.Value)
                        c.cliente = Convert.ToString(r["cliente"]);

                    if (r["cups"] != System.DBNull.Value)
                        c.cups = Convert.ToString(r["cups"]);

                    if (r["producto"] != System.DBNull.Value)
                        c.producto = Convert.ToString(r["producto"]);

                    if (r["fecha_inicio"] != System.DBNull.Value)
                        c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

                    d.Add(c.id, c);

                }
                db.CloseConnection();

                return d;
            }catch(Exception e)
            {
                ficheroLog.AddError("AdicionalesSharepoint.CargaAdicionalesProcesados: " + e.Message);
                hay_error = true;
                return null;
            }
        }

        private void DescargaArchivoExcelSharepoint()
        {
            string destino = "";
            utilidades.Credenciales credenciales = new utilidades.Credenciales();

            try
            {
                ficheroLog.Add("Inicio descarga de formulario Excel");
                ficheroLog.Add("================================");
                ficheroLog.Add("");

                sharePoint.userName = credenciales.server_user;
                sharePoint.password = credenciales.server_password;

                ficheroLog.Add("Conectando con el sitio: " + param.GetValue("Adicionales_siteURL", DateTime.Now, DateTime.Now));
                Console.WriteLine("Conectando con el sitio: " + param.GetValue("Adicionales_siteURL", DateTime.Now, DateTime.Now));

                sharePoint.siteURL = param.GetValue("Adicionales_siteURL", DateTime.Now, DateTime.Now);
                destino = param.GetValue("Adicionales_Carpeta_Descarga", DateTime.Now, DateTime.Now);

                ficheroLog.Add("Descargando el archivo: " + param.GetValue("Adicionales_Excel", DateTime.Now, DateTime.Now));
                Console.WriteLine("Descargando el archivo: " + param.GetValue("Adicionales_Excel", DateTime.Now, DateTime.Now));

                sharePoint.fileURL = param.GetValue("AZUCARERA_fileURL", DateTime.Now, DateTime.Now)
                    + param.GetValue("Adicionales_Excel", DateTime.Now, DateTime.Now);
                sharePoint.destination = destino +
                    param.GetValue("Adicionales_Excel", DateTime.Now, DateTime.Now);
                sharePoint.Download();


            }
            catch(Exception e)
            {
                hay_error = true;
                ficheroLog.AddError("AdicionalesSharepoint.DescargaArchivoExcelSharepoint: " + e.Message);
            }
        }

        private void BorrarContenidoDirectorio(string directorio)
        {
            string[] listaArchivos;
            FileInfo file;
            try
            {
                listaArchivos = Directory.GetFiles(directorio);
                for (int i = 0; i < listaArchivos.Count(); i++)
                {
                    file = new FileInfo(listaArchivos[i]);
                    file.Delete();
                }
                hay_error = false;
            }
            catch (Exception e)
            {
                hay_error = true;
            }

        }

        private Dictionary<int, EndesaEntity.contratacion.gas.AdicionalSharepoint> LeeExcel()
        {
            Dictionary<int, EndesaEntity.contratacion.gas.AdicionalSharepoint> d =
                new Dictionary<int, EndesaEntity.contratacion.gas.AdicionalSharepoint>();
                
                        

            int c = 1;
            int f = 1;
            string hora = "";
            DateTime diaLectura = new DateTime();            

            try
            {
                diaLectura = DateTime.Now.AddDays(1); // Hoy no Mañaaaaana                


                ficheroLog.Add("Inicio lectura de archivo Excel");
                ficheroLog.Add("================================");
                ficheroLog.Add("");

                Console.WriteLine("");

                string[] listaArchivosExcel = Directory.GetFiles(param.GetValue("Adicionales_Carpeta_Descarga", DateTime.Now, DateTime.Now), "*.xlsx");
                for (int x = 0; x < listaArchivosExcel.Length; x++)
                {

                    FileInfo file = new FileInfo(listaArchivosExcel[x]);

                    ficheroLog.Add("Mirando dentro del archivo: " + file.Name);
                    Console.WriteLine("Mirando dentro del archivo: " + file.Name);

                    FileStream fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    ExcelPackage excelPackage = new ExcelPackage(fs);
                    var workSheet = excelPackage.Workbook.Worksheets.First();

                    //sol = new EndesaEntity.contratacion.gas.Solicitud();
                    //sol.cups = param.GetValue(file.Name, DateTime.Now, DateTime.Now);                    

                    f = 1; // Porque la primera fila es la cabecera
                    for (int i = 0; i < 100000; i++)
                    {
                        f++;
                        c = 1;

                        EndesaEntity.contratacion.gas.AdicionalSharepoint e = new EndesaEntity.contratacion.gas.AdicionalSharepoint();
                        e.id = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                        e.hora_de_inicio = Convert.ToDateTime(workSheet.Cells[f, c].Value); c++;
                        e.hora_de_finalizacion = Convert.ToDateTime(workSheet.Cells[f, c].Value); c++;
                        e.correo_electronico = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                        e.nombre = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                        e.cliente = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                        e.cups = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                        e.producto = Convert.ToString(workSheet.Cells[f, c].Value).ToUpper(); c++;
                        e.fecha_inicio = Convert.ToDateTime(workSheet.Cells[f, c].Value); c++;

                        if(e.producto == "INTRADIARIO")
                        {
                            hora = Convert.ToString(workSheet.Cells[f, c].Value);
                            e.fecha_inicio.AddHours(Convert.ToInt32(hora.Substring(0, 2)));
                            e.fecha_inicio.AddMinutes(Convert.ToInt32(hora.Substring(3, 2)));
                        }

                        EndesaEntity.contratacion.gas.AdicionalSharepoint o;
                        if (!dic_adicionales_procesados.TryGetValue(e.id, out o))                        
                            if (!d.TryGetValue(e.id, out o))
                                d.Add(e.id, e);
                    }

                    ficheroLog.Add("Cerrando el archivo: " + fs.Name);

                    fs.Close();
                    fs = null;
                    excelPackage.Dispose();
                    excelPackage = null;
                }

                ficheroLog.Add("Fin lectura de archivo Excel");
                ficheroLog.Add("================================");
                ficheroLog.Add("");

                hay_error = false;
                return d;

            }
            catch (Exception e)
            {
                hay_error = true;
                ficheroLog.AddError("Azucarera.LeeExcel: " + e.Message);
                return null;
            }
        }

        private void GeneraXML(Dictionary<int, EndesaEntity.contratacion.gas.AdicionalSharepoint> dic)
        {
            cnmc.CNMC cnmc = new cnmc.CNMC();
            cnmc.XML formato_xml = new cnmc.XML();
            //EndesaBusiness.utilidades.ZIP zip = new utilidades.ZIP();
            EndesaBusiness.utilidades.ZipUnZip zip = new utilidades.ZipUnZip();
            contratacion.gestionATRGas.Distribuidoras distribuidoras = new Distribuidoras(true);
            int secuencial;
            string fileName = "";
            string fechaHora = "";
            string destinycompany = "";

            try
            {
                if (dic.Count() > 0)
                {
                    foreach (KeyValuePair<int, EndesaEntity.contratacion.gas.AdicionalSharepoint> p in dic)
                    {
                        secuencial = Convert.ToInt32(param.GetValue("secuencial_solicitud", DateTime.Now, DateTime.Now)) + 1;

                        EndesaEntity.cnmc.XML_A1_43 xml_a1_43 = new EndesaEntity.cnmc.XML_A1_43();

                        xml_a1_43.comreferencenum = param.GetValue("prefijo_solicitud", DateTime.Now, DateTime.Now)
                            + DateTime.Now.ToString("yyyy") + secuencial.ToString().PadLeft(4, '0');

                        xml_a1_43.destinycompany =
                            distribuidoras.Codigo_XML_CNMC_Distribuidora(param.GetValue("nombre_distribuidora_azucarera", DateTime.Now, DateTime.Now));
                        destinycompany = xml_a1_43.destinycompany;
                        xml_a1_43.dispatchingcompany = "0007";
                        xml_a1_43.documentnum = param.GetValue("AZUCARERA_NIF", DateTime.Now, DateTime.Now);
                        xml_a1_43.cups = p.Value.cups;
                        //xml_a1_43.productstartdate = p.detalle[0].fecha_inicio;
                        //xml_a1_43.producttype = cnmc.Codigo_Tipo_Producto(p.detalle[0].producto);
                        //xml_a1_43.producttolltype = cnmc.Codigo_Tipo_Peaje(p.detalle[0].qd * 330);
                        //xml_a1_43.productqd = p.Value.qd;

                        fechaHora = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
                        fileName = param.GetValue("AZUCARERA_Carpeta_Descarga", DateTime.Now, DateTime.Now)
                            + param.GetValue("prefijo_archivo_a1_43", DateTime.Now, DateTime.Now)
                            + xml_a1_43.destinycompany + "_"
                            + fechaHora
                            + ".xml";
                        FileInfo file = new FileInfo(fileName);

                        formato_xml.CreaXML_A1_43(file, xml_a1_43);
                        //GuardaNumSecuencialTemporal(secuencial);
                    }

                    // Comprimimos los XML y los enviamos por mail
                    fechaHora = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    fileName = param.GetValue("AZUCARERA_Carpeta_Descarga", DateTime.Now, DateTime.Now)
                            + param.GetValue("prefijo_archivo_a1_43", DateTime.Now, DateTime.Now)
                            + destinycompany + "_"
                            + fechaHora
                            + ".zip";
                    FileInfo archivo = new FileInfo(fileName);
                    //zip.ComprimirVarios(param.GetValue("AZUCARERA_Carpeta_Descarga", DateTime.Now, DateTime.Now),
                    //    ".*\\.(xml)$", archivo.FullName);

                    zip.ComprimirVarios(param.GetValue("AZUCARERA_Carpeta_Descarga") + "*.xml", archivo.FullName);

                    //GeneraMail(fileName);
                }

            }
            catch (Exception e)
            {
                hay_error = true;
                ficheroLog.AddError("Azucarera.GeneraXML: " + e.Message);
            }
        }

        private void GuardaDatosBBDD(Dictionary<int, EndesaEntity.contratacion.gas.AdicionalSharepoint> dic)
        {
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            MySQLDB db;
            MySqlCommand command;
            int x = 0;
            int i = 0;


            try
            {

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand("delete from atrgas_solicitudes_sharepoint_tmp", db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                foreach (KeyValuePair<int,EndesaEntity.contratacion.gas.AdicionalSharepoint> p in dic)
                {
                    x++;
                    i++;
                    if (firstOnly)
                    {
                        sb.Append("replace into atrgas_solicitudes_sharepoint_tmp (");
                        sb.Append("id, hora_de_inicio, hora_de_finalizacion, correo_electronico, nombre,");
                        sb.Append("cliente, cups, producto, fecha_inicio, hora_inicio, generado_XML) values ");
                        firstOnly = false;
                    }

                    #region Campos
                    sb.Append("(").Append(p.Value.id).Append(",");
                    sb.Append("'").Append(p.Value.hora_de_inicio.ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    sb.Append("'").Append(p.Value.hora_de_finalizacion.ToString("yyyy-MM-dd HH:mm:ss")).Append("',");                    
                    sb.Append("'").Append(p.Value.correo_electronico).Append("',");
                    sb.Append("'").Append(p.Value.nombre).Append("',");
                    sb.Append("'").Append(p.Value.cliente).Append("',");
                    sb.Append("'").Append(p.Value.cups).Append("',");
                    sb.Append("'").Append(p.Value.producto).Append("',");
                    sb.Append("'").Append(p.Value.fecha_inicio.ToString("yyyy-MM-dd HH:mm:ss")).Append("',");                                        
                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                    #endregion

                    if (x == 100)
                    {
                        Console.WriteLine("Guardando " + i + " registros...");
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
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
                    Console.Write("");
                    Console.Write("Guardando " + x + " registros...");
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    x = 0;
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine("GuardaDatosBBDDExtraccionFormulas --> " + e.Message);
                ficheroLog.AddError("GuardaDatosBBDDExtraccionFormulas --> " + e.Message);
            }
        }
    }
}
