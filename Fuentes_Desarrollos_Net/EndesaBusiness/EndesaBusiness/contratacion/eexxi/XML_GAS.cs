using EndesaBusiness.servidores;
using Microsoft.Graph;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Renci.SshNet.Common;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Telegram.Bot.Types;

namespace EndesaBusiness.contratacion.eexxi
{
    public class XML_GAS
    {
        Dictionary<string, List<EndesaEntity.xml.CNMC_t_cabecera>> dic_cabeceras;
        Dictionary<string, int> dic_archivos;
        EndesaBusiness.utilidades.Param p;
        public XML_GAS() 
        {
            //dic_archivos = CargaArchivos();
            p = new EndesaBusiness.utilidades.Param("gas_eexxi_xml_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
        }

        private Dictionary<string, int> CargaArchivos()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            Dictionary<string, int> d =
                new Dictionary<string, int>();

            try
            {
                strSql = "select id, filename"
                    + " from gas_eexxi_xml_files order by id";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                }
                db.CloseConnection();

                return d;

            }catch(Exception ex) 
            {
                return null;
            }

                
        }
        public bool Guarda_XML(string archivo, EndesaEntity.cnmc.gas.Proceso_A15_50 xml)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int i = 0;
            int id = 0;
            try
            {
                id = GetID(archivo);
                                
                sb.Append("REPLACE INTO gas_eexxi_xml_heading");
                sb.Append(" (id, dispatchingcode, dispatchingcompany, destinycompany,");
                sb.Append(" communicationsdate, communicationshour, processcode, messagetype,");
                sb.Append(" created_by, created_date) values ");

                id = GetID(archivo);

                sb.Append("(").Append(id).Append(",");
                sb.Append("'").Append(xml.heading.dispatchingcode).Append("',");
                sb.Append("'").Append(xml.heading.dispatchingcompany).Append("',");
                sb.Append("'").Append(xml.heading.destinycompany).Append("',");
                sb.Append("'").Append(xml.heading.communicationsdate.ToString("yyyy-MM-dd")).Append("',");
                sb.Append("'").Append(xml.heading.communicationshour.ToString("HH:mm:ss")).Append("',");
                sb.Append("'").Append(xml.heading.processcode).Append("',");
                sb.Append("'").Append(xml.heading.messagetype).Append("',");                
                sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");               
                

                firstOnly = true;
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                sb = null;
                sb = new StringBuilder();
                i = 0;

                
                Guarda_Body(id, archivo, xml);
                Guarda_Address(id, archivo, xml);
                

                //Guarda_DatosSolicitud(id, xml);
                //Guarda_Contacto(id, xml);
                //Guarda_Cliente(id, xml);

                return true;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                        "Importación ficheros XML",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                return false;
            }
        }

        private bool Guarda_Body(int id, string archivo, EndesaEntity.cnmc.gas.Proceso_A15_50 xml)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int i = 0;

            try
            {
                sb.Append("REPLACE INTO gas_eexxi_xml_body");
                sb.Append(" (id, reqcode, responsedate, responsehour, cups, atrcode,");
                sb.Append(" transfereffectivedate, enservicio, tolltype, telemetering,");
                sb.Append(" finalclientyearlyconsumption, gasusetype, netsituation, result,");
                sb.Append(" resultdesc, nationality, documenttype, documentnum, titulartype, firstname,");
                sb.Append(" familyname1, familyname2, telephone1, email, canonircperiodicity, ");
                sb.Append(" lastinspectionsdate, lastinspectionsresult, statusps, readingtype,");
                sb.Append(" created_by, created_date) values ");


                sb.Append("(").Append(id).Append(",");
                sb.Append("'").Append(xml.a1550.reqcode).Append("',");
                sb.Append("'").Append(xml.a1550.responsedate.ToString("yyyy-MM-dd")).Append("',");
                sb.Append("'").Append(xml.a1550.responsehour.ToString("HH:mm:ss")).Append("',");
                sb.Append("'").Append(xml.a1550.cups).Append("',");
                sb.Append("'").Append(xml.a1550.atrcode).Append("',");
                sb.Append("'").Append(xml.a1550.transfereffectivedate.ToString("yyyy-MM-dd")).Append("',");
                sb.Append("'").Append(xml.a1550.enservicio).Append("',");
                sb.Append("'").Append(xml.a1550.tolltype).Append("',");
                sb.Append("'").Append(xml.a1550.telemetering).Append("',");
                sb.Append("'").Append(xml.a1550.finalclientyearlyconsumption).Append("',");
                sb.Append("'").Append(xml.a1550.gasusetype).Append("',");
                sb.Append("'").Append(xml.a1550.netsituation).Append("',");
                sb.Append("'").Append(xml.a1550.result).Append("',");
                sb.Append("'").Append(xml.a1550.resultdesc).Append("',");
                sb.Append("'").Append(xml.a1550.nationality).Append("',");
                sb.Append("'").Append(xml.a1550.documenttype).Append("',");
                sb.Append("'").Append(xml.a1550.documentnum).Append("',");

                if (xml.a1550.titulartype != null)
                    sb.Append("'").Append(xml.a1550.titulartype).Append("',");
                else
                    sb.Append("null,");


                if (xml.a1550.firstname != null)
                    sb.Append("'").Append(utilidades.FuncionesTexto.RT(xml.a1550.firstname)).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1550.familyname1 != null)
                    sb.Append("'").Append(utilidades.FuncionesTexto.RT(xml.a1550.familyname1)).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1550.familyname2 != null)
                    sb.Append("'").Append(utilidades.FuncionesTexto.RT(xml.a1550.familyname2)).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1550.telephone1 != null)
                    sb.Append("'").Append(xml.a1550.telephone1).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1550.email != null)
                    sb.Append("'").Append(xml.a1550.email).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1550.canonircperiodicity != null)
                    sb.Append("'").Append(xml.a1550.canonircperiodicity).Append("',");
                else
                    sb.Append("null,");

                sb.Append("'").Append(xml.a1550.lastinspectionsdate.ToString("yyyy-MM-dd")).Append("',");
                sb.Append("'").Append(xml.a1550.lastinspectionsresult).Append("',");
                sb.Append("'").Append(xml.a1550.StatusPS).Append("',");

                if (xml.a1550.readingtype != null)
                    sb.Append("'").Append(xml.a1550.readingtype).Append("',");
                else
                    sb.Append("null,");

                sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                firstOnly = true;
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                sb = null;
                sb = new StringBuilder();
                i = 0;


                return true;
            }catch (Exception ex) 
            {
                MessageBox.Show(ex.Message,
                       "Guarda_Body",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Error);

                return false;
            }


        }

        private bool Guarda_Address(int id, string archivo, EndesaEntity.cnmc.gas.Proceso_A15_50 xml)
        {

            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int i = 0;

            try
            {
                sb.Append("REPLACE INTO gas_eexxi_xml_addressps");
                sb.Append(" (id, province, city, zipcode, streettype, street,");
                sb.Append(" streetnumber, portal, floor, door,");                                
                sb.Append(" created_by, created_date) values ");


                sb.Append("(").Append(id).Append(",");
                sb.Append("'").Append(xml.a1550.AddressPS.province).Append("',");
                sb.Append("'").Append(xml.a1550.AddressPS.city).Append("',");
                sb.Append("'").Append(xml.a1550.AddressPS.zipcode).Append("',");
                sb.Append("'").Append(xml.a1550.AddressPS.street.streettype).Append("',");

                if (xml.a1550.AddressPS.street.street != null)
                    sb.Append("'").Append(utilidades.FuncionesTexto.RT(xml.a1550.AddressPS.street.street)).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1550.AddressPS.street.streetnumber == null)
                    sb.Append("'").Append(xml.a1550.AddressPS.street.streetnumber).Append("',");
                else
                    sb.Append("null,");
                
                sb.Append("null,"); // portal

                if (xml.a1550.AddressPS.street.floor == null)
                    sb.Append("'").Append(xml.a1550.AddressPS.street.floor).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1550.AddressPS.street.door == null)
                    sb.Append("'").Append(xml.a1550.AddressPS.street.door).Append("',");
                else
                    sb.Append("null,");              


                sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                firstOnly = true;
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                sb = null;
                sb = new StringBuilder();
                i = 0;

                return true;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message,
                       "Guarda_Address",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Error);

                return false;
            }
            
        }

        private int GetID(string archivo)
        {
            int id = 0;

            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            strSql = "select id from gas_eexxi_xml_files"
                + " where filename = '" + archivo + "'";
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            if (r.Read())
                id = Convert.ToInt32(r.GetString("id"));
            db.CloseConnection();

            if(id == 0)
            {
                id = GetLastID();
                strSql = "REPLACE INTO gas_eexxi_xml_files"
                    + " (id, filename, created_by, created_date) values"
                    + " (" + id + ","
                    + "'" + archivo + "',"
                    + "'" + System.Environment.UserName.ToUpper() + "',"
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }

            return id;           

            
        }
        private int GetLastID()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            int id = 0;

            strSql = "select max(id) id from gas_eexxi_xml_files";
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
                if(r["id"] != System.DBNull.Value)
                    id = Convert.ToInt32(r["id"]);

            db.CloseConnection();

            return id + 1;

        }

        public void InformeExcel(List<EndesaEntity.contratacion.gas.Informe_GAS_XML> lista)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                ExportExcel(save.FileName,  lista);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }
        }

        private void ExportExcel(string rutaFichero, List<EndesaEntity.contratacion.gas.Informe_GAS_XML> lista)
        {

            int numDias = 0;
            DateTime sfd = new DateTime();
            DateTime sfh = new DateTime();

            int f = 1;
            int c = 1;


            bool firstOnly = true;

            FileInfo file = new FileInfo(rutaFichero);

            if (file.Exists)
                file.Delete();

            // Ruta de la plantilla 
            FileInfo plantillaExcel = new FileInfo(System.Environment.CurrentDirectory + p.GetValue("plantilla_informe"));



            ExcelPackage excelPackage;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            excelPackage = new ExcelPackage(plantillaExcel);
            var workSheet = excelPackage.Workbook.Worksheets["Datos"];
            

            f = 2;
            foreach(EndesaEntity.contratacion.gas.Informe_GAS_XML p in lista)
            {
                c = 1;
                f++;

                workSheet.Cells[f, c].Value = p.archivo; c++;
                workSheet.Cells[f, c].Value = p.cups_xml; c++;
                workSheet.Cells[f, c].Value = p.tarifa_xml; c++;
                workSheet.Cells[f, c].Value = p.fecha_inicio_xml;
                workSheet.Cells[f, c].Style.Numberformat.Format =
                    DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                c++;

                //if(p.sistema != null)
                //    workSheet.Cells[f, c].Value = p.sistema;
                //c++;

                //if (p.cups_sistema != null)
                //    workSheet.Cells[f, c].Value = p.cups_sistema;
                //c++;

                //if (p.tarifa_sistema != null)
                //    workSheet.Cells[f, c].Value = p.tarifa_sistema;
                //c++;

                //if(p.fecha_inicio_sistema != DateTime.MinValue)
                //{
                //    workSheet.Cells[f, c].Value = p.fecha_inicio_sistema;
                //    workSheet.Cells[f, c].Style.Numberformat.Format =
                //        DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                //}
                //c++;

            }



            var allCells = workSheet.Cells[1, 1, f, c];
            allCells.AutoFitColumns();

            excelPackage.SaveAs(file);
            excelPackage = null;
            
        }

            
        public List<EndesaEntity.contratacion.gas.Informe_GAS_XML> Datos()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            bool comprobar_rosetta = false;

            List<EndesaEntity.contratacion.gas.Informe_GAS_XML> lista =
                new List<EndesaEntity.contratacion.gas.Informe_GAS_XML>();

            Dictionary<string, string> dic_cups20 = new Dictionary<string, string>();

            try
            {
                strSql = "SELECT f.created_date AS fecha_importacion," 
                    + " c.cups, c.transfereffectivedate AS fecha_inicio,"
                    + " c.tolltype AS tarifa,"
                    + " f.filename AS archivo"
                    + " FROM gas_eexxi_xml_files f"
                    + " INNER join gas_eexxi_xml_body c ON"
                    + " c.id = f.id";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.gas.Informe_GAS_XML c =
                        new EndesaEntity.contratacion.gas.Informe_GAS_XML();

                    if (r["archivo"] != System.DBNull.Value)
                        c.archivo = r["archivo"].ToString();

                    if (r["fecha_importacion"] != System.DBNull.Value)
                        c.fecha_importacion = Convert.ToDateTime(r["fecha_importacion"]);

                    if (r["cups"] != System.DBNull.Value)
                        c.cups_xml = r["cups"].ToString();

                    if (r["fecha_inicio"] != System.DBNull.Value)
                        c.fecha_inicio_xml = Convert.ToDateTime(r["fecha_inicio"]);

                    if (r["tarifa"] != System.DBNull.Value)
                        c.tarifa_xml = r["tarifa"].ToString();


                    lista.Add(c);

                    

                }                    
                db.CloseConnection();

                // Comprobación SIGAME

                Dictionary<string, List<EndesaEntity.sigame.ContratoGas>>  dic = Inventario_SIGAME_EEXXI();
                foreach (EndesaEntity.contratacion.gas.Informe_GAS_XML p in lista)
                {
                    List<EndesaEntity.sigame.ContratoGas> o;
                    if (dic.TryGetValue(p.cups_xml, out o))
                    {
                        foreach (EndesaEntity.sigame.ContratoGas pp in o)
                        {
                            if (pp.fecha_inicio == p.fecha_inicio_xml)
                            {
                                p.sistema = "SIGAME";
                                p.cups_sistema = pp.cupsree;
                                p.fecha_inicio_sistema = pp.fecha_inicio;
                                p.tarifa_sistema = pp.tarifa;
                            }
                        }

                        if(p.sistema == null)
                        {
                            string oo;
                            if (!dic_cups20.TryGetValue(p.cups_xml, out oo))
                                dic_cups20.Add(p.cups_xml, p.cups_xml);
                        }
                    }
                    else
                    {
                        string aa;
                        if (!dic_cups20.TryGetValue(p.cups_xml, out aa))
                            dic_cups20.Add(p.cups_xml, p.cups_xml);

                    }                      
                }

                // Comprobación Rosetta
                if (comprobar_rosetta)
                {
                    Redshift.ContratosRosetta rosetta = new Redshift.ContratosRosetta(dic_cups20.Values.ToList());

                    foreach (EndesaEntity.contratacion.gas.Informe_GAS_XML p in lista)
                    {
                        if (p.sistema == null)
                        {
                            List<EndesaEntity.contratacion.PS_AT_Tabla> i;
                            if (rosetta.dic.TryGetValue(p.cups_xml, out i))
                            {
                                foreach (EndesaEntity.contratacion.PS_AT_Tabla g in i)
                                {
                                    if (g.fecha_alta_contrato == p.fecha_inicio_xml)
                                    {
                                        p.sistema = "ROSETTA";
                                        p.cups_sistema = g.cups20;
                                        p.fecha_inicio_sistema = g.fecha_alta_contrato;
                                        p.tarifa_sistema = g.tarifa;
                                    }
                                }
                            }
                        }
                    }
                }
                return lista;
            }
            catch(Exception ex) 
            {
                return null;
            }

            
        }

        public bool XML_Valido(string fileName)
        {
            FileInfo file = new FileInfo(fileName);
            EndesaEntity.contratacion.xxi.XML_Datos xml = new EndesaEntity.contratacion.xxi.XML_Datos();

            XmlTextReader r;
            try
            {

                r = new XmlTextReader(fileName);
                while (r.Read())
                {

                }

                return true;
            }
            catch(Exception ex)
            {
                return false;
            }

        }
        
        private Dictionary<string, List<EndesaEntity.sigame.ContratoGas>> Inventario_SIGAME_EEXXI()
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql = "";

            Dictionary<string, List<EndesaEntity.sigame.ContratoGas>> dic =
                new Dictionary<string, List<EndesaEntity.sigame.ContratoGas>>();

            strSql = "SELECT DISTINCT T_SGM_G_PS.ID_PS, T_SGM_G_PS.DE_PS, T_SGM_M_CLIENTES.CD_CIF,T_SGM_M_CLIENTES.DE_NOMBRE_CLIENTE,"
                + " T_SGM_G_CONTRATOS_PS.ID_CTO_PS, T_SGM_G_CONTRATOS_PS.FH_INICIO_REAL, T_SGM_G_CONTRATOS_PS.FH_FIN_REAL,"
                + " T_SGM_G_CONTRATOS_PS.FH_FIN_PREVISTA, T_SGM_G_CONTRATOS_PS.ID_ESTADO_CTO, T_SGM_P_TIPO_TARIFA.DE_TIPO_TARIFA," 
                + " T_SGM_G_PS.CD_CUPS, T_SGM_M_GESTORES.DE_GESTOR, DE_NOMBRE_MUNICIPIO, T_SGM_P_PROVINCIAS.DE_PROVINCIA,T_SGM_G_PS.CD_PAIS," 
                + " T_SGM_P_DISTRIBUIDORES.DE_DISTRIBUIDORES, T_SGM_M_RED_DISTRIBUCION.DE_RED_DISTRIBUCION, (T_SGM_G_PS.NM_SUMA_CONSUMOS_MEN) / 12 AS PROMEDIO,"
                + " T_SGM_G_PS.NM_ENERGIA_DIARIA_KWH, T_SGM_G_CONTRATOS_PS.FH_ULT_ACTUALIZACION,T_SGM_G_CONTRATOS_PS.CD_COMERCIALIZADORA"
                + " FROM T_SGM_G_PS"
                + " LEFT JOIN T_SGM_G_CONTRATOS_PS ON T_SGM_G_PS.ID_PS = T_SGM_G_CONTRATOS_PS.ID_PS"
                + " LEFT JOIN T_SGM_M_CLIENTES ON T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE LEFT JOIN T_SGM_M_GESTORES ON T_SGM_M_CLIENTES.ID_GESTOR = T_SGM_M_GESTORES.ID_GESTOR"
                + " LEFT JOIN T_SGM_P_TIPO_TARIFA ON T_SGM_G_PS.ID_TIPO_TARIFA = T_SGM_P_TIPO_TARIFA.ID_TIPO_TARIFA LEFT JOIN T_SGM_P_DISTRIBUIDORES ON T_SGM_G_PS.ID_DISTRIBUIDORES = T_SGM_P_DISTRIBUIDORES.ID_DISTRIBUIDORES"
                + " LEFT JOIN T_SGM_M_RED_DISTRIBUCION ON  T_SGM_G_PS.ID_RED_DISTRIBUCION = T_SGM_M_RED_DISTRIBUCION.ID_RED_DISTRIBUCION"
                + " LEFT JOIN T_SGM_P_PROVINCIAS ON T_SGM_G_PS.ID_PROVINCIA = T_SGM_P_PROVINCIAS.ID_PROVINCIA AND T_SGM_G_PS.CD_PAIS = T_SGM_P_PROVINCIAS.CD_PAIS"
                + " LEFT JOIN T_SGM_P_MUNICIPIOS ON  T_SGM_G_PS.CD_MUNICIPIO = T_SGM_P_MUNICIPIOS.CD_MUNICIPIO AND T_SGM_G_PS.CD_PAIS = T_SGM_P_MUNICIPIOS.CD_PAIS"
                + " WHERE T_SGM_G_CONTRATOS_PS.CD_COMERCIALIZADORA = 2";

            db = new SQLServer();
            command = new SqlCommand(strSql, db.con);

            SqlDataAdapter da = new SqlDataAdapter(command);
            r = command.ExecuteReader();

            while (r.Read())
            {

                EndesaEntity.sigame.ContratoGas c = new EndesaEntity.sigame.ContratoGas();
                if (r["ID_PS"] != System.DBNull.Value)
                {
                    c.id_ps = Convert.ToInt32(r["ID_PS"]);                    
                }

                if (r["ID_CTO_PS"] != System.DBNull.Value)
                    c.id_cto_ps = Convert.ToInt32(r["ID_CTO_PS"]);

                if (r["DE_PS"] != System.DBNull.Value)
                    c.descripcion_ps = Convert.ToString(r["DE_PS"]);

                if (r["DE_NOMBRE_CLIENTE"] != System.DBNull.Value)
                    c.nombre_cliente = Convert.ToString(r["DE_NOMBRE_CLIENTE"]);

                //if (r["NM_ENERGIA_DIARIA_KWH"] != System.DBNull.Value)
                //    c.qd = Convert.ToDouble(r["NM_ENERGIA_DIARIA_KWH"]);

                c.nif = Convert.ToString(r["CD_CIF"]);

                if (r["FH_INICIO_REAL"] != System.DBNull.Value)
                    c.fecha_inicio = Convert.ToDateTime(r["FH_INICIO_REAL"]);
                else
                    c.fecha_inicio = DateTime.MinValue;

                if (r["FH_FIN_REAL"] != System.DBNull.Value)
                    c.fecha_fin = Convert.ToDateTime(r["FH_FIN_REAL"]);
                else
                    c.fecha_fin = new DateTime(4999, 12, 31);

                if (r["ID_ESTADO_CTO"] != System.DBNull.Value)
                    c.id_estado_contrato = Convert.ToInt32(r["ID_ESTADO_CTO"]);
                else
                    c.id_estado_contrato = 0;

                if (r["DE_TIPO_TARIFA"] != System.DBNull.Value)
                {
                    c.tarifa = Convert.ToString(r["DE_TIPO_TARIFA"]);
                }


                if (r["CD_CUPS"] != System.DBNull.Value)
                {

                    c.cupsree = Convert.ToString(r["CD_CUPS"]).Trim();
                    if (c.cupsree.Length > 2)
                    {
                        c.pais = c.cupsree.Substring(0, 2) == "PT" ? "Portugal" : "España";
                        c.es_cisterna = false;

                        if (c.tarifa != null)
                        {
                            int n;
                            if (int.TryParse(c.tarifa.Substring(0, 1), out n))
                                c.grupo_presion = Convert.ToInt32(c.tarifa.Substring(0, 1));
                        }


                    }
                    else
                    {
                        c.es_cisterna = true;
                        c.cupsree = "Cisterna_" + c.id_ps;
                        c.pais = "España";
                    }

                }
                else
                {
                    c.cupsree = "Cisterna_" + c.id_ps;
                    c.es_cisterna = true;
                    c.pais = "España";
                }


                if (r["DE_DISTRIBUIDORES"] != System.DBNull.Value)
                    c.distribuidora = r["DE_DISTRIBUIDORES"].ToString().Trim();

                
                List<EndesaEntity.sigame.ContratoGas> o;
                if (!dic.TryGetValue(c.cupsree, out o))
                {
                    o = new List<EndesaEntity.sigame.ContratoGas>();
                    o.Add(c);
                    dic.Add(c.cupsree, o);
                }
                else
                    o.Add(c);
                        
                

            }
            db.CloseConnection();

            return dic;
        }
                


    }
}
