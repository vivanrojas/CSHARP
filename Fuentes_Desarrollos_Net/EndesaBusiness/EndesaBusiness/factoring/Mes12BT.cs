using EndesaBusiness.servidores;
using Microsoft.Graph.TermStore;
using Microsoft.Graph;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using Renci.SshNet.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Telegram.Bot.Types;
using Org.BouncyCastle.Utilities.Collections;
using EndesaEntity;
using static EndesaBusiness.medida.Kee_Extraccion_Formulas;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using EndesaEntity.factoring;
using static Microsoft.Exchange.WebServices.Data.SearchFilter;

namespace EndesaBusiness.factoring
{
    public class Mes12BT
    {
        public bool es_xlsb { get; set; }
        public string fichero_excel_mes12 { get; set; }

        EndesaBusiness.utilidades.Param p;
        // clave == hoja excel
        Dictionary<string, List<EndesaEntity.factoring.Mes12_BT>> dic;
        Dictionary<string, string> dic_referencias;
        Dictionary<string, string> dic_referencias_nif;
        logs.Log ficheroLog;
        public List<string> lista_adjudicaciones;
        public List<string> lista_previsiones;
        public List<EndesaEntity.factoring.DatosExcel> datos_excel;

        public Mes12BT()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "Mes12BT");
            p = new EndesaBusiness.utilidades.Param("mes12_param", MySQLDB.Esquemas.FAC);
            lista_adjudicaciones = Carga_Codigos_Adjudicaciones();
            lista_previsiones = Carga_Codigos_Previsiones();
            datos_excel = Carga_Resumen_Datos_Excel();
        }


        public List<string> Carga_Codigos_Adjudicaciones()
        {
            List<string> l = new List<string>();

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";


            strSql = "select factoring from mes12_adjudicaciones group by factoring order by factoring desc";               
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["factoring"] != System.DBNull.Value)
                    l.Add(r["factoring"].ToString());
            }
            db.CloseConnection();

            return l;
        }

        public List<string> Carga_Codigos_Previsiones()
        {
            List<string> l = new List<string>();

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";


            strSql = "select factoring from mes12_previsiones group by factoring order by factoring desc";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["factoring"] != System.DBNull.Value)
                    l.Add(r["factoring"].ToString());
            }
            db.CloseConnection();

            return l;
        }

        private List<EndesaEntity.factoring.DatosExcel> Carga_Resumen_Datos_Excel()
        {
            List<EndesaEntity.factoring.DatosExcel> l = new List<EndesaEntity.factoring.DatosExcel>();

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";


            strSql = "SELECT c.HOJA, COUNT(*) AS total_registros FROM mes12_excel c GROUP BY c.HOJA";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.factoring.DatosExcel c = new EndesaEntity.factoring.DatosExcel();
                if (r["HOJA"] != System.DBNull.Value)
                    c.hoja = r["HOJA"].ToString();
                if (r["total_registros"] != System.DBNull.Value)
                    c.registros = Convert.ToInt32(r["total_registros"]);

                l.Add(c);
            }
            db.CloseConnection();

            return l;
        }



        public void CargaExcel(string fichero)
        {

            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "replace into mes12_excel_hist"
                + " select * from mes12_excel";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "delete from mes12_excel";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


            dic = ProcesaExcel(fichero);
            if(dic != null)
            {
                GuardaExcel_enBBDD(dic);
            //    Complementa_CUPS();
            }
                

            

            //RellenaInventario(dic);
            //Proceso(mensual);
            //GeneraExcelResultados(fichero);
        }

        private Dictionary<string, List<EndesaEntity.factoring.Mes12_BT>> ProcesaExcel(string fichero)
        {
            int c = 1;
            int f = 1;
            int total_hojas_excel = 0;
            bool firstOnly = true;
            string cabecera = "";
            FileStream fs;
            ExcelPackage excelPackage;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;

            Dictionary<string, List<EndesaEntity.factoring.Mes12_BT>> d =
                new Dictionary<string, List<EndesaEntity.factoring.Mes12_BT>>();

            List<EndesaEntity.factoring.Mes12_BT> lista;

            try
            {
                ficheroLog.Add("Cargando archivo: " + fichero);

                fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                excelPackage = new ExcelPackage(fs);
                
                total_hojas_excel = excelPackage.Workbook.Worksheets.Count();
                var workSheet = excelPackage.Workbook.Worksheets.First();

                for (int hoja = 0; hoja < (total_hojas_excel - 1); hoja++)
                {
                    workSheet = excelPackage.Workbook.Worksheets[hoja];
                    firstOnly = true;

                    if (workSheet.Cells[1, 1].Value == null)
                    {
                        ficheroLog.Add("Total filas " + workSheet.Name + ": " + String.Format("{0:N0}", f));
                        break;
                    }
                        


                    lista = new List<EndesaEntity.factoring.Mes12_BT>();

                    if (workSheet.Name.Contains("FACTORIZABLES"))
                    {
                        f = 1; // Porque la primera fila es la cabecera
                        for (int i = 1; i < 1048576; i++)
                        {
                            c = 1;
                            f++;

                            if (workSheet.Cells[f, 1].Value == null)
                                break;

                            if (workSheet.Cells[f, 1].Value.ToString() == "")
                                break;

                            EndesaEntity.factoring.Mes12_BT t = new EndesaEntity.factoring.Mes12_BT();

                            t.hoja = workSheet.Name;
                            t.cnifdnic = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.dapersoc = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.dprovinc = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.sociedad = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.creferen = Convert.ToInt64(workSheet.Cells[f, c].Value); c++;
                            t.secfactu = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.toblfrac = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.tclisegm = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.tcliente = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.tformpag = Convert.ToString(workSheet.Cells[f, c].Value).Trim(); c++;

                            t.iobligac = Convert.ToDouble(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.ffactura = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.fptacobr = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.flimpago = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.fcobobli = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            t.testimpg = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.cmedpago = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.ffactori = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            t.cfactura = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.entrada = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.factorizadas = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.procesos_concursales = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.lista_negra = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.menores_1000 = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.menores_50000 = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            //t.ddd_vencida = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.csegmerc = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            t.tfactura = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.ctarifa = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.ttension = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.cd_ind_cef_ope = Convert.ToString(workSheet.Cells[f, c].Value).Trim(); c++;

                            if (workSheet.Cells[f, c].Value != null)
                            {
                                t.cd_pun_notif = Convert.ToString(workSheet.Cells[f, c].Value).Trim();
                                if (t.cd_pun_notif.Length > 20)
                                    t.cups20 = t.cd_pun_notif.Substring(0, 20);
                                else
                                    t.cups20 = t.cd_pun_notif;
                            }
                                
                            c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.tipo_cliente = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.tipo_ne = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            List<EndesaEntity.factoring.Mes12_BT> o;
                            if (!d.TryGetValue(t.hoja, out o))
                            {
                                o = new List<EndesaEntity.factoring.Mes12_BT>();
                                o.Add(t);
                                d.Add(t.hoja, o);
                            }
                            else
                                o.Add(t);


                        }
                    }


                    if (workSheet.Name == "FACTURAS VENCIDAS AAPP")
                    {
                        f = 1; // Porque la primera fila es la cabecera
                        for (int i = 1; i < 1048576; i++)
                        {
                            c = 1;
                            f++;

                            if (workSheet.Cells[f, 1].Value == null)
                                break;

                            if (workSheet.Cells[f, 1].Value.ToString() == "")
                                break;

                            EndesaEntity.factoring.Mes12_BT t = new EndesaEntity.factoring.Mes12_BT();

                            t.hoja = workSheet.Name;
                            t.cnifdnic = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.dapersoc = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.dprovinc = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.sociedad = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.creferen = Convert.ToInt64(workSheet.Cells[f, c].Value); c++;
                            t.secfactu = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.toblfrac = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.tclisegm = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.tcliente = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.tformpag = Convert.ToString(workSheet.Cells[f, c].Value).Trim(); c++;

                            t.iobligac = Convert.ToDouble(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.ffactura = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.fptacobr = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.flimpago = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.fcobobli = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            t.testimpg = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.destimpg = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            t.cmedpago = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.ffactori = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            t.cfactura = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                            t.entrada = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.factorizadas = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.procesos_concursales = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.lista_negra = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.menores_1000 = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.menores_50000 = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                            t.ddd_vencida = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.tipo_cliente = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.csegmerc = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            t.tfactura = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.ctarifa = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.ttension = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.cd_ind_cef_ope = Convert.ToString(workSheet.Cells[f, c].Value).Trim(); c++;

                            if (workSheet.Cells[f, c].Value != null)
                                t.cd_pun_notif = Convert.ToString(workSheet.Cells[f, c].Value).Trim(); c++;
                           

                            List<EndesaEntity.factoring.Mes12_BT> o;
                            if (!d.TryGetValue(t.hoja, out o))
                            {
                                o = new List<EndesaEntity.factoring.Mes12_BT>();
                                o.Add(t);
                                d.Add(t.hoja, o);
                            }
                            else
                                o.Add(t);


                        }
                    }

                }
                fs = null;
                excelPackage = null;


                #region Buscamos facturas sin CUPS
                Dictionary<string, string> dic_facturas_sin_cups = new Dictionary<string, string>();
                // Buscamos las facturas sin CUPS
                foreach(KeyValuePair<string, List<EndesaEntity.factoring.Mes12_BT>> x in d)
                {
                    foreach(EndesaEntity.factoring.Mes12_BT p in x.Value) 
                    {
                        if (p.cd_pun_notif == null || p.cd_pun_notif.Trim() == "")
                        {
                            string o;
                            if (!dic_facturas_sin_cups.TryGetValue(p.cfactura, out o))
                                dic_facturas_sin_cups.Add(p.cfactura, p.cfactura);
                        }
                            
                    }
                }
                #endregion

                Dictionary<string, string> dic_facturas_cups = Relaccion_Facturas_CUPS(dic_facturas_sin_cups);


                #region Asignamos las cups encontrados al Excel
                foreach (KeyValuePair<string, List<EndesaEntity.factoring.Mes12_BT>> x in d)
                {
                    foreach (EndesaEntity.factoring.Mes12_BT p in x.Value)
                    {
                        if (p.cd_pun_notif == null || p.cd_pun_notif.Trim() == "")
                        {
                            string o;
                            if (dic_facturas_cups.TryGetValue(p.cfactura, out o))
                            {
                                p.cd_pun_notif = o;
                                if (o.Length > 20)
                                    p.cups20 = o.Substring(1, 20);
                                else
                                    p.cups20 = o;
                            }
                                
                        }

                    }
                }
                #endregion


                return d;
            }
            catch(Exception ex)
            {
                fs = null;
                excelPackage = null;
                ficheroLog.AddError("ProcesaExcel: " + ex.Message);
                MessageBox.Show(ex.Message,
                    "ProcesaExcel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return null;
            }
        }


        private Dictionary<string, string> Relaccion_Facturas_CUPS(Dictionary<string, string> dic_facturas)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            bool firstOnly = true;
            string cfactura = "";
            string cupsree = "";
            StringBuilder sb = new StringBuilder();


            Dictionary<string, string> d = new Dictionary<string, string>();

            try
            {
                // FACTURAS SCE
                sb.Append("Select CFACTURA, CUPSREE from fo where");
                sb.Append(" CFACTURA in (");

                foreach(KeyValuePair<string, string> p in dic_facturas)
                {
                    if (firstOnly)
                    {
                        sb.Append("'").Append(p.Key).Append("'");
                        firstOnly = false;
                    }
                    else
                    {
                        sb.Append(",'").Append(p.Key).Append("'");
                    }
                }

                sb.Append(")");
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();
                while (r.Read()) 
                {
                    cfactura = r["CFACTURA"].ToString();
                    if (r["CUPSREE"] != System.DBNull.Value)
                        cupsree = r["CUPSREE"].ToString();

                    string o;
                    if (!d.TryGetValue(cfactura, out o))
                    {
                        d.Add(cfactura, cupsree);
                    }
                }
                db.CloseConnection();



                // FACTURAS SIGAME
                sb = null;
                sb = new StringBuilder();
                firstOnly = true;
                sb.Append("SELECT CD_NFACTURA_REALES_PS, CUPS22 FROM fo_s WHERE");
                sb.Append(" CD_NFACTURA_REALES_PS in (");

                foreach (KeyValuePair<string, string> p in dic_facturas)
                {
                    
                    if (firstOnly)
                    {
                        sb.Append("'").Append(p.Key).Append("'");
                        firstOnly = false;
                    }
                    else
                    {
                        sb.Append(",'").Append(p.Key).Append("'");
                    }
                }

                sb.Append(")");
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["CD_NFACTURA_REALES_PS"] != System.DBNull.Value)
                        cfactura = r["CD_NFACTURA_REALES_PS"].ToString();

                    if (r["CUPS22"] != System.DBNull.Value)
                        cupsree = r["CUPS22"].ToString();

                    string o;
                    if (!d.TryGetValue(cfactura, out o))
                    {
                        d.Add(cfactura, cupsree);
                    }
                }
                db.CloseConnection();


                return d;

            }
            catch(Exception ex)
            {
                return null;
            }
        }

        private void GuardaExcel_enBBDD(Dictionary<string, List<EndesaEntity.factoring.Mes12_BT>> d)
        {

            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            int numReg = 0;
            int totalReg = 0;
            bool firstOnly = true;
            string strSql = "";
            string hoja = "";
            

            try
            {

                // BorradoTabla_Mes12_Excel();

                foreach (KeyValuePair<string, List<EndesaEntity.factoring.Mes12_BT>> kv in d)
                { 

                #region Hojas Iguales
                foreach (EndesaEntity.factoring.Mes12_BT p in kv.Value)
                {
                    numReg++;

                    hoja = p.hoja;
                    if (hoja.Contains("FACTORIZABLES"))
                    {
                        if (firstOnly)
                        {
                            sb.Append("REPLACE INTO mes12_excel");
                            sb.Append(" (VERSION, HOJA, CNIFDNIC, DAPERSOC, DPROVINC, SOCIEDAD, CREFEREN, SECFACTU, TOBLFRAC, TCLISEGM, TCLIENTE,");
                            sb.Append(" TFORMPAG, IOBLIGAC, FFACTURA, FPTACOBR, FLIMPAGO, FCOBOBLI, TESTIMPG, CMEDPAGO, FFACTORI, CFACTURA, ENTRADA,");
                            sb.Append(" FACTORIZADAS, PROCESOS_CONCURSALES, LISTA_NEGRA, MENORES_1000, MENORES_50000, CSEGMERC, TFACTURA, CTARIFA,");
                            sb.Append(" TTENSION, CD_IND_CEF_OPE, CD_PUN_NOTIF, CUPS20, TIPO_CLIENTE, TIPO_NE, TIPO, REFERENCIA, created_by, created_date) VALUES ");

                            firstOnly = false;
                        }

                        #region Campos
                        sb.Append("('").Append(DateTime.Now.Date.ToString("yyyyMMdd")).Append("',");
                        sb.Append("'").Append(p.hoja).Append("',");
                        sb.Append("'").Append(p.cnifdnic).Append("',");
                        sb.Append("'").Append(utilidades.FuncionesTexto.RT(p.dapersoc)).Append("',");

                        if (p.dprovinc != null)
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(p.dprovinc)).Append("',");
                        else
                            sb.Append("null,");

                        if (p.sociedad != null)
                            sb.Append("'").Append(p.sociedad).Append("',");
                        else
                            sb.Append("null,");

                        if (p.creferen != 0)
                            sb.Append("'").Append(p.creferen).Append("',");
                        else
                            sb.Append("null,");

                        if (p.secfactu != 0)
                            sb.Append("'").Append(p.secfactu).Append("',");
                        else
                            sb.Append("null,");

                        if (p.toblfrac != null)
                            sb.Append("'").Append(p.toblfrac).Append("',");
                        else
                            sb.Append("null,");

                        if (p.tclisegm != null)
                            sb.Append("'").Append(p.tclisegm).Append("',");
                        else
                            sb.Append("null,");

                        if (p.tcliente != null)
                            sb.Append("'").Append(p.tcliente).Append("',");
                        else
                            sb.Append("null,");

                        if (p.tformpag != null)
                            sb.Append("'").Append(p.tformpag).Append("',");
                        else
                            sb.Append("null,");

                        sb.Append(p.iobligac.ToString().Replace(",", ".")).Append(",");

                        if (p.ffactura != null)
                            sb.Append("'").Append(p.ffactura).Append("',");
                        else
                            sb.Append("null,");

                        if (p.fptacobr != null)
                            sb.Append("'").Append(p.fptacobr).Append("',");
                        else
                            sb.Append("null,");

                        if (p.flimpago != null)
                            sb.Append("'").Append(p.flimpago).Append("',");
                        else
                            sb.Append("null,");

                        if (p.fcobobli != null)
                            sb.Append("'").Append(p.fcobobli).Append("',");
                        else
                            sb.Append("null,");

                        if (p.testimpg != 0)
                            sb.Append(p.testimpg).Append(",");
                        else
                            sb.Append("null,");

                        if (p.cmedpago != 0)
                            sb.Append(p.cmedpago).Append(",");
                        else
                            sb.Append("null,");

                        if (p.ffactori != null)
                            sb.Append("'").Append(p.ffactori).Append("',");
                        else
                            sb.Append("null,");

                        if (p.cfactura != null)
                            sb.Append("'").Append(p.cfactura).Append("',");
                        else
                            sb.Append("null,");

                        sb.Append(p.entrada).Append(",");
                        sb.Append(p.factorizadas).Append(",");
                        sb.Append(p.procesos_concursales).Append(",");
                        sb.Append(p.lista_negra).Append(",");
                        sb.Append(p.menores_1000).Append(",");
                        sb.Append(p.menores_50000).Append(",");




                        if (p.csegmerc != null)
                            sb.Append("'").Append(p.csegmerc).Append("',");
                        else
                            sb.Append("null,");

                        sb.Append(p.tfactura).Append(",");

                        if (p.ctarifa != null)
                            sb.Append("'").Append(p.ctarifa).Append("',");
                        else
                            sb.Append("null,");

                        if (p.ttension != null)
                            sb.Append("'").Append(p.ttension).Append("',");
                        else
                            sb.Append("null,");

                        if (p.cd_ind_cef_ope != null)
                            sb.Append("'").Append(p.cd_ind_cef_ope).Append("',");
                        else
                            sb.Append("null,");

                        if (p.cd_pun_notif != null)
                            sb.Append("'").Append(p.cd_pun_notif).Append("',");
                        else
                            sb.Append("null,");

                        if (p.cups20 != null)
                            sb.Append("'").Append(p.cups20).Append("',");
                        else
                            sb.Append("null,");

                            if (p.tipo_cliente != null)
                            sb.Append("'").Append(p.tipo_cliente).Append("',");
                        else
                            sb.Append("null,");

                        if (p.tipo_ne != null)
                            sb.Append("'").Append(p.tipo_ne).Append("',");
                        else
                            sb.Append("null,");

                        sb.Append("null,"); // TIPO
                        sb.Append("null,"); // REFERENCIA

                        sb.Append("'").Append(System.Environment.UserName).Append("',");
                        sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                        #endregion


                        if (numReg == 250)
                        {
                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.FAC);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            numReg = 0;
                        }
                    }

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
                sb = new StringBuilder();
                numReg = 0;
            }

            #endregion

            foreach (KeyValuePair<string, List<EndesaEntity.factoring.Mes12_BT>> kv in d)
                {
                    

                    foreach (EndesaEntity.factoring.Mes12_BT p in kv.Value)
                    {
                        

                        hoja = p.hoja;                        

                        if (hoja == "FACTURAS VENCIDAS AAPP")
                        {
                            numReg++;

                            if (firstOnly)
                            {
                                sb.Append("REPLACE INTO mes12_excel");
                                sb.Append(" (VERSION, HOJA, CNIFDNIC, DAPERSOC, DPROVINC, SOCIEDAD, CREFEREN, SECFACTU, TOBLFRAC, TCLISEGM, TCLIENTE,");
                                sb.Append(" TFORMPAG, IOBLIGAC, FFACTURA, FPTACOBR, FLIMPAGO, FCOBOBLI, TESTIMPG, DESTIMPG, CMEDPAGO, FFACTORI, CFACTURA, ENTRADA,");
                                sb.Append(" FACTORIZADAS, PROCESOS_CONCURSALES, LISTA_NEGRA, MENORES_1000, MENORES_50000, DDD_VENCIDA, TIPO_CLIENTE, CSEGMERC, TFACTURA, CTARIFA,");
                                sb.Append(" TTENSION, CD_IND_CEF_OPE, CD_PUN_NOTIF, CUPS20, TIPO, REFERENCIA, created_by, created_date) VALUES ");

                                firstOnly = false;
                            }

                            #region Campos
                            sb.Append("('").Append(DateTime.Now.Date.ToString("yyyyMMdd")).Append("',");
                            sb.Append("'").Append(p.hoja).Append("',");
                            sb.Append("'").Append(p.cnifdnic).Append("',");
                            sb.Append("'").Append(utilidades.FuncionesTexto.RT(p.dapersoc)).Append("',");

                            if (p.dprovinc != null)
                                sb.Append("'").Append(utilidades.FuncionesTexto.RT(p.dprovinc)).Append("',");
                            else
                                sb.Append("null,");

                            if (p.sociedad != null)
                                sb.Append("'").Append(p.sociedad).Append("',");
                            else
                                sb.Append("null,");

                            if (p.creferen != 0)
                                sb.Append("'").Append(p.creferen).Append("',");
                            else
                                sb.Append("null,");

                            if (p.secfactu != 0)
                                sb.Append("'").Append(p.secfactu).Append("',");
                            else
                                sb.Append("null,");

                            if (p.toblfrac != null)
                                sb.Append("'").Append(p.toblfrac).Append("',");
                            else
                                sb.Append("null,");

                            if (p.tclisegm != null)
                                sb.Append("'").Append(p.tclisegm).Append("',");
                            else
                                sb.Append("null,");

                            if (p.tcliente != null)
                                sb.Append("'").Append(p.tcliente).Append("',");
                            else
                                sb.Append("null,");

                            if (p.tformpag != null)
                                sb.Append("'").Append(p.tformpag).Append("',");
                            else
                                sb.Append("null,");

                            sb.Append(p.iobligac.ToString().Replace(",", ".")).Append(",");

                            if (p.ffactura != null)
                                sb.Append("'").Append(p.ffactura).Append("',");
                            else
                                sb.Append("null,");

                            if (p.fptacobr != null)
                                sb.Append("'").Append(p.fptacobr).Append("',");
                            else
                                sb.Append("null,");

                            if (p.flimpago != null)
                                sb.Append("'").Append(p.flimpago).Append("',");
                            else
                                sb.Append("null,");

                            if (p.fcobobli != null)
                                sb.Append("'").Append(p.fcobobli).Append("',");
                            else
                                sb.Append("null,");

                            if (p.testimpg != 0)
                                sb.Append(p.testimpg).Append(",");
                            else
                                sb.Append("null,");

                            if (p.destimpg != null)
                                sb.Append("'").Append(p.destimpg).Append("',");
                            else
                                sb.Append("null,");

                            if (p.cmedpago != 0)
                                sb.Append(p.cmedpago).Append(",");
                            else
                                sb.Append("null,");

                            if (p.ffactori != null)
                                sb.Append("'").Append(p.ffactori).Append("',");
                            else
                                sb.Append("null,");

                            if (p.cfactura != null)
                                sb.Append("'").Append(p.cfactura).Append("',");
                            else
                                sb.Append("null,");

                            sb.Append(p.entrada).Append(",");
                            sb.Append(p.factorizadas).Append(",");
                            sb.Append(p.procesos_concursales).Append(",");
                            sb.Append(p.lista_negra).Append(",");
                            sb.Append(p.menores_1000).Append(",");
                            sb.Append(p.menores_50000).Append(",");
                            sb.Append(p.ddd_vencida).Append(",");

                            if (p.tipo_cliente != null)
                                sb.Append("'").Append(p.tipo_cliente).Append("',");
                            else
                                sb.Append("null,");

                            if (p.csegmerc != null)
                                sb.Append("'").Append(p.csegmerc).Append("',");
                            else
                                sb.Append("null,");

                            sb.Append(p.tfactura).Append(",");

                            if (p.ctarifa != null)
                                sb.Append("'").Append(p.ctarifa).Append("',");
                            else
                                sb.Append("null,");

                            if (p.ttension != null)
                                sb.Append("'").Append(p.ttension).Append("',");
                            else
                                sb.Append("null,");

                            if (p.cd_ind_cef_ope != null)
                                sb.Append("'").Append(p.cd_ind_cef_ope).Append("',");
                            else
                                sb.Append("null,");

                            if (p.cd_pun_notif != null)
                                sb.Append("'").Append(p.cd_pun_notif).Append("',");
                            else
                                sb.Append("null,");

                            if (p.cups20 != null)
                                sb.Append("'").Append(p.cups20).Append("',");
                            else
                                sb.Append("null,");

                            sb.Append("null,"); // TIPO
                            sb.Append("null,"); // REFERENCIA

                            sb.Append("'").Append(System.Environment.UserName).Append("',");
                            sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                            #endregion


                            if (numReg == 250)
                            {
                                firstOnly = true;
                                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                                command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                                command.ExecuteNonQuery();
                                db.CloseConnection();
                                sb = null;
                                sb = new StringBuilder();
                                numReg = 0;
                            }
                        }
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
                    sb = new StringBuilder();
                    numReg = 0;
                }


            }
            catch(Exception ex)
            {
                ficheroLog.AddError("GuardaExcel_enBBDD: " + ex.Message);

                MessageBox.Show(ex.Message,
                   "GuardaExcel_enBBDD",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }

        public void Complementa_CUPS()
        {

            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "UPDATE mes12_excel SET cups20 = CD_PUN_NOTIF"
                + " WHERE LENGTH(CD_PUN_NOTIF) > 1 AND  LENGTH(CD_PUN_NOTIF) < 20";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "UPDATE mes12_excel SET cups20 = SUBSTR(CD_PUN_NOTIF, 1, 20)"
                + " WHERE CD_PUN_NOTIF <> ''";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "UPDATE mes12_excel SET cups20 = SUBSTR(CD_PUN_NOTIF, 1, 20)"
                + " WHERE CD_PUN_NOTIF is not null";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


            strSql = "UPDATE mes12_excel t"
                + " INNER JOIN fo f ON"
                + " f.CNIFDNIC = t.CNIFDNIC and"
                + " f.CFACTURA = t.CFACTURA"
                + " SET CD_PUN_NOTIF = f.CUPSREE,"
                + " t.CUPS20 = substr(f.CUPSREE,1,20),"
                + " t.last_update_by = '" + System.Environment.UserName.ToUpper() + "'"
                + " WHERE t.CUPS20 = ''";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "UPDATE mes12_excel t"
                + " INNER JOIN fo f ON"
                + " f.CNIFDNIC = t.CNIFDNIC and"
                + " f.CFACTURA = t.CFACTURA"
                + " SET CD_PUN_NOTIF = f.CUPSREE,"
                + " t.CUPS20 = substr(f.CUPSREE,1,20),"
                + " t.last_update_by = '" + System.Environment.UserName.ToUpper() + "'"
                + " WHERE t.CUPS20 is NULL";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


            // Para las facturas de GAS
            strSql = "UPDATE mes12_excel t"
                + " INNER JOIN fo_s f ON"
                + " f.CIFNIF = t.CNIFDNIC and"
                + " f.CD_NFACTURA_REALES_PS = t.CFACTURA"
                + " SET t.CD_PUN_NOTIF = f.CUPS22,"
                + " t.CUPS20 = f.CUPS22,"
                + " t.last_update_by = '" + System.Environment.UserName.ToUpper() + "'"
                + " WHERE t.CUPS20 = ''";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "UPDATE mes12_excel t"
                + " INNER JOIN fo_s f ON"
                + " f.CIFNIF = t.CNIFDNIC and"
                + " f.CD_NFACTURA_REALES_PS = t.CFACTURA"
                + " SET t.CD_PUN_NOTIF = f.CUPS22,"
                + " t.CUPS20 = f.CUPS22,"
                + " t.last_update_by = '" + System.Environment.UserName.ToUpper() + "'"
                + " WHERE t.CUPS20 is NULL";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


        }        
        public void CruzaConPrevisiones(string factoring, string hoja)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "UPDATE mes12_excel t"
                + " INNER JOIN mes12_previsiones aa ON"
                + " aa.CUPSREE = t.CUPS20"
                + " SET t.TIPO = 'EN PREVISIÓN MES 13',"
                + " t.REFERENCIA = aa.REFERENCIA"
                + " WHERE t.HOJA = '" + hoja + "' AND"
                + " aa.factoring = '" + factoring + "' AND"
                + " t.CUPS20 <> ''";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            //strSql = "UPDATE mes12_excel t"
            //    + " INNER JOIN mes12_previsiones aa ON"
            //    + " aa.CUPSREE = t.CUPS20"
            //    + " SET t.TIPO = 'EN PREVISIÓN MES 13',"
            //    + " t.REFERENCIA = aa.REFERENCIA"
            //    + " WHERE t.HOJA = '" + hoja + "' AND"
            //    + " aa.factoring = '" + factoring + "' AND"
            //    + " t.CUPS20 is null";
            //ficheroLog.Add(strSql);
            //db = new MySQLDB(MySQLDB.Esquemas.FAC);
            //command = new MySqlCommand(strSql, db.con);
            //command.ExecuteNonQuery();
            //db.CloseConnection();

            strSql = "UPDATE mes12_excel t"
                + " INNER JOIN mes12_previsiones aa ON"
                + " aa.NIF = t.CNIFDNIC"
                + " SET t.TIPO = 'EN PREVISIÓN MES 13',"
                + " t.REFERENCIA = aa.REFERENCIA"
                + " WHERE t.HOJA = '" + hoja + "' AND"
                + " aa.factoring = '" + factoring + "' AND"
                + " t.CUPS20 = '' AND"
                + " t.TFACTURA = 6";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "UPDATE mes12_excel t"
                + " INNER JOIN mes12_previsiones aa ON"
                + " aa.NIF = t.CNIFDNIC"
                + " SET t.TIPO = 'EN PREVISIÓN MES 13',"
                + " t.REFERENCIA = aa.REFERENCIA"
                + " WHERE t.HOJA = '" + hoja + "' AND"
                + " aa.factoring = '" + factoring + "' AND"
                + " t.CUPS20 is null AND"
                + " t.TFACTURA = 6";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }
        public void CruzaConAdjudicaciones(string hoja)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "UPDATE mes12_excel t"
                + " INNER JOIN mes12_adjudicaciones aa ON"
                + " aa.CUPSREE = t.CUPS20"
                + " SET t.TIPO = 'EN ADJUDICACIÓN MES 13',"
                + " t.REFERENCIA = aa.REFERENCIA"
                + " WHERE t.HOJA = '" + hoja + "' AND"                
                + " t.CUPS20 <> ''";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            //strSql = "UPDATE mes12_excel t"
            //    + " INNER JOIN mes12_previsiones aa ON"
            //    + " aa.CUPSREE = t.CUPS20"
            //    + " SET t.TIPO = 'EN PREVISIÓN MES 13',"
            //    + " t.REFERENCIA = aa.REFERENCIA"
            //    + " WHERE t.HOJA = '" + hoja + "' AND"
            //    + " aa.factoring = '" + factoring + "' AND"
            //    + " t.CUPS20 is null";
            //ficheroLog.Add(strSql);
            //db = new MySQLDB(MySQLDB.Esquemas.FAC);
            //command = new MySqlCommand(strSql, db.con);
            //command.ExecuteNonQuery();
            //db.CloseConnection();

            strSql = "UPDATE mes12_excel t"
                + " INNER JOIN mes12_adjudicaciones aa ON"
                + " aa.NIF = t.CNIFDNIC"
                + " SET t.TIPO = 'EN ADJUDICACIÓN MES 13',"
                + " t.REFERENCIA = aa.REFERENCIA"
                + " WHERE t.HOJA = '" + hoja + "' AND"                
                + " t.CUPS20 = '' AND"
                + " t.TFACTURA in (5,6)";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "UPDATE mes12_excel t"
                + " INNER JOIN mes12_adjudicaciones aa ON"
                + " aa.NIF = t.CNIFDNIC"
                + " SET t.TIPO = 'EN ADJUDICACIÓN MES 13',"
                + " t.REFERENCIA = aa.REFERENCIA"
                + " WHERE t.HOJA = '" + hoja + "' AND"                
                + " t.CUPS20 is null AND"
                + " t.TFACTURA in (5,6)";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            // Marcamos todos los NIF de la adjudicación en Mes12
            // para asegurarnos que no nos dejamos ninguno.

            strSql = "UPDATE mes12_excel t"
                + " INNER JOIN mes12_adjudicaciones aa ON"
                + " aa.NIF = t.CNIFDNIC"
                + " SET t.NIF_EN_MES13 = 'NIF EN MES13'"
                + " WHERE t.HOJA = '" + hoja + "'";                
                
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }
        public void ResetExcel()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "UPDATE mes12_excel t"
                + " SET t.TIPO = null,"
                + " t.REFERENCIA = null,"
                + " t.NIF_EN_MES13 = null";                
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }
        public void GeneraRespuesta()
        {
            int c = 1;
            int f = 1;
            int total_hojas_excel = 0;
            bool firstOnly = true;
            bool firstOnlyNIF = true;
            string cabecera = "";
            string factura = "";
            string referencia = "";
            string referencia_nif = "";

            FileInfo fileInfo = new FileInfo(p.GetValue("nombre_archivo"));
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            total_hojas_excel = excelPackage.Workbook.Worksheets.Count();
            var workSheet = excelPackage.Workbook.Worksheets.First();

            FileInfo ficheroDestino = new FileInfo(fileInfo.FullName.Replace(".xlsx","_RESPUESTA.xlsx"));
            this.fichero_excel_mes12 = ficheroDestino.FullName;
            dic_referencias = CargaDatosExcelConReferencias();
            dic_referencias_nif = CargaDatosExcelConNIFAdjudicaciones();

            for (int hoja = 0; hoja < (total_hojas_excel - 1); hoja++)
            {
                workSheet = excelPackage.Workbook.Worksheets[hoja];
                firstOnly = true;
                firstOnlyNIF = true;    

                if (workSheet.Name.Contains("FACTORIZABLES") && !workSheet.Name.Contains("NO "))
                {
                    f = 1; // Porque la primera fila es la cabecera
                    for (int i = 1; i < 1048576; i++)
                    {
                        c = 1;
                        f++;

                        if (workSheet.Cells[f, 1].Value == null)
                            break;

                        if (workSheet.Cells[f, 1].Value.ToString() == "")
                            break;

                        if (workSheet.Cells[f, 19].Value != null)
                        {
                            factura = workSheet.Cells[f, 19].Value.ToString();
                            referencia = GetReferencia(factura);
                            if (referencia != null)
                            {
                                if (firstOnly)
                                {
                                    workSheet.Cells[1, 34].Value = "TIPO";
                                    workSheet.Cells[1, 35].Value = "REFERENCIA";
                                    firstOnly = false;                                    
                                }


                                workSheet.Cells[f, 34].Value = "EN ADJUDICACIÓN MES 13";
                                workSheet.Cells[f, 35].Value = referencia;
                            }

                            referencia_nif = GetReferenciaNIF(factura);
                            if (referencia_nif != null)
                            {
                                if (firstOnlyNIF)
                                {
                                
                                    workSheet.Cells[1, 36].Value = "NIF en MES13";
                                    firstOnlyNIF = false;
                                }
                                
                                workSheet.Cells[f, 36].Value = referencia_nif;
                            }

                        }
                    }
                }

                if (workSheet.Name == "FACTURAS VENCIDAS AAPP")
                {
                    f = 1; // Porque la primera fila es la cabecera
                    for (int i = 1; i < 1048576; i++)
                    {
                        c = 1;
                        f++;

                        if (workSheet.Cells[f, 1].Value == null)
                            break;

                        if (workSheet.Cells[f, 1].Value.ToString() == "")
                            break;

                        if (workSheet.Cells[f, 20].Value != null)
                        {
                            factura = workSheet.Cells[f, 20].Value.ToString();
                            referencia = GetReferencia(factura);
                            if (referencia != null)
                            {
                                if (firstOnly)
                                {
                                    workSheet.Cells[1, 35].Value = "TIPO";
                                    workSheet.Cells[1, 36].Value = "REFERENCIA";
                                    firstOnly = false;
                                }

                                workSheet.Cells[f, 35].Value = "EN ADJUDICACIÓN MES 13";
                                workSheet.Cells[f, 36].Value = referencia;
                            }

                            referencia_nif = GetReferenciaNIF(factura);
                            if (referencia_nif != null)
                            {
                                if (firstOnlyNIF)
                                {
                                    workSheet.Cells[1, 37].Value = "NIF en MES13";
                                    firstOnlyNIF = false;
                                }

                                workSheet.Cells[f, 37].Value = referencia_nif;
                            }
                        }

                    }
                }
            }

            excelPackage.SaveAs(ficheroDestino);
        }

        public void ImportarAdjudicacion(string fichero)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            int linea = 0;
            bool firstOnly = true;
            bool firstOnlyFactoring = true;
            string cabecera = "";
            string strSql = "";
            string factoring = "";
            List<EndesaEntity.factoring.Seguimiento> lista = new List<EndesaEntity.factoring.Seguimiento>();


            StringBuilder sb = new StringBuilder();
            MySQLDB db;
            MySqlCommand command;

            try
            {

                // FileInfo file = new FileInfo(fichero);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();

                f = 1; // Porque la primera fila es la cabecera
                for (int i = 0; i < 100000; i++)
                {
                    linea = 1;
                    c = 1;

                    if (firstOnly)
                    {
                        for (int w = 1; w < 11; w++)
                            cabecera += workSheet.Cells[1, w].Value.ToString();


                        if (!EstructuraCorrecta(cabecera))
                        {
                            // this.hayError = true;
                            // this.descripcion_error = "La estructura del archivo excel no es la correcta.";

                            MessageBox.Show("La estructura del archivo no es la correcta!!!"
                                + System.Environment.NewLine
                                + "Se esperaba la cabecera: ENTIDAD LN EMPRESA TITULAR NIF CLIENTE CCOUNIPS CUPSREE REFERENCIA SEC CONTROL",
                                "Estructura Incorrecta Excel",
                                MessageBoxButtons.OK,
                            MessageBoxIcon.Error);

                            break;
                        }

                        firstOnly = false;
                    }

                    f++;

                    cabecera = "";
                    if (workSheet.Cells[f, 1].Value != null &&
                       workSheet.Cells[f, 2].Value != null &&
                       workSheet.Cells[f, 3].Value != null)
                    {

                        EndesaEntity.factoring.Seguimiento s = new EndesaEntity.factoring.Seguimiento();

                        // SEGMENTO ENTIDAD LN EMPRESATITULAR NIF CLIENTE CCOUNIPS CUPS REEREFERENCIA SEC
                        // s.segmento = workSheet.Cells[f, c].Value.ToString(); c++;
                        s.entidad = workSheet.Cells[f, c].Value.ToString(); c++;
                        s.linea_negocio = workSheet.Cells[f, c].Value.ToString().Substring(0, 1) == "L" ? "LUZ" : "GAS"; c++;
                        s.empresa_titular = workSheet.Cells[f, c].Value.ToString(); c++;
                        s.nif = workSheet.Cells[f, c].Value.ToString(); c++;
                        s.nombre_cliente = workSheet.Cells[f, c].Value.ToString(); c++;

                        if (workSheet.Cells[f, c].Value != null)
                            if (workSheet.Cells[f, c].Value.ToString().Trim() != "")
                                s.cups13 = workSheet.Cells[f, c].Value.ToString();
                        c++;

                        if (workSheet.Cells[f, c].Value != null)
                        {
                            if (workSheet.Cells[f, c].Value.ToString().Trim() != "")
                                s.cups20 = workSheet.Cells[f, c].Value.ToString();
                        }
                        else if (s.cups13 != "")
                            s.cups20 = s.cups13;

                        c++;

                        s.referencia = workSheet.Cells[f, c].Value.ToString(); c++;

                        if (firstOnlyFactoring)
                        {
                            factoring = s.referencia.Substring(0, 6);
                            firstOnlyFactoring = false;
                        }
                        
                        s.secuencial = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;                        
                        s.control = workSheet.Cells[f, c].Value.ToString(); c++;
                        s.estimacion_importe = Convert.ToDouble(workSheet.Cells[f, c].Value); c++;
                        s.estimacion_base = Convert.ToDouble(workSheet.Cells[f, c].Value); c++;
                        s.estimacion_impuestos = Convert.ToDouble(workSheet.Cells[f, c].Value); c++;
                        s.diaf = Convert.ToDateTime(workSheet.Cells[f, c].Value.ToString()); c++;
                        s.diav = Convert.ToDateTime(workSheet.Cells[f, c].Value.ToString()); c++;

                        lista.Add(s);

                    }

                }

                fs = null;
                excelPackage = null;

                if (lista.Count> 0 && UltimaAdjudicacion() != factoring) 
                {
                    strSql = "DELETE FROM mes12_adjudicaciones";
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                }

                firstOnly = true;
                id = 0;
                for (int i = 0; i < lista.Count; i++)
                {
                    id++;
                    if (firstOnly)
                    {
                        sb.Append("replace into mes12_adjudicaciones (factoring, nif, cliente,");
                        sb.Append(" cupsree, referencia, created_by, created_date) values ");
                        firstOnly = false;
                    }



                    sb.Append("('").Append(factoring).Append("',");                    
                    sb.Append("'").Append(lista[i].nif).Append("',");
                    sb.Append("'").Append(lista[i].nombre_cliente).Append("',");                   

                    if (lista[i].cups20 != null)
                        sb.Append("'").Append(lista[i].cups20).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append(lista[i].referencia).Append("',");
                    sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");


                    if (id == 250)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        id = 0;
                    }

                }

                if (id > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    id = 0;
                }

               



               


                MessageBox.Show("Importación finalizada."
                          + System.Environment.NewLine
                          + System.Environment.NewLine
                          + System.Environment.NewLine,
                    "Proceso Mes12 BT",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception e)
            {
                MessageBox.Show("Error en la línea " + linea + " --> " + e.Message,
                  "Error en la importación de adjudicaciones",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        public string UltimaAdjudicacion()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string factoring = "";

            strSql = "SELECT min(d.factoring) AS factoring FROM mes12_adjudicaciones d";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if(r["factoring"] != System.DBNull.Value)
                    factoring = r["factoring"].ToString();
            }
            db.CloseConnection();
            return factoring;
        }

        public void ImportarPrevision(string fichero)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            int linea = 0;
            bool firstOnly = true;
            string cabecera = "";
            
            List<EndesaEntity.factoring.Seguimiento> lista = new List<EndesaEntity.factoring.Seguimiento>();

            FileStream fs;
            ExcelPackage excelPackage;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;

            StringBuilder sb = new StringBuilder();
            MySQLDB db;
            MySqlCommand command;

            int total_hojas_excel;
            try
            {

                // FileInfo file = new FileInfo(fichero);
                fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);                
                excelPackage = new ExcelPackage(fs);
                total_hojas_excel = excelPackage.Workbook.Worksheets.Count();
                var workSheet = excelPackage.Workbook.Worksheets.First();

                for (int hoja = 0; hoja < 2; hoja++)
                {
                    workSheet = excelPackage.Workbook.Worksheets[hoja];
                    firstOnly = true;

                    f = 1; // Porque la primera fila es la cabecera
                    for (int i = 0; i < 100000; i++)
                    {
                        linea = 1;
                        c = 1;

                        if (firstOnly)
                        {
                            for (int w = 1; w < 8; w++)
                                cabecera += workSheet.Cells[1, w].Value.ToString();


                            if (!EstructuraCorrectaPrevision(cabecera))
                            {
                                // this.hayError = true;
                                // this.descripcion_error = "La estructura del archivo excel no es la correcta.";

                                MessageBox.Show("La estructura del archivo no es la correcta",
                                    "Estructura Incorrecta Excel",
                                    MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                                break;
                            }

                            firstOnly = false;
                        }

                        f++;

                        cabecera = "";
                        if (workSheet.Cells[f, 1].Value != null &&
                           workSheet.Cells[f, 2].Value != null &&
                           workSheet.Cells[f, 3].Value != null)
                        {

                            EndesaEntity.factoring.Seguimiento s = new EndesaEntity.factoring.Seguimiento();
                            s.factoring = workSheet.Cells[f, 7].Value.ToString().Substring(0, 6);
                            s.nif = workSheet.Cells[f, 3].Value.ToString();
                            s.nombre_cliente = workSheet.Cells[f, 4].Value.ToString(); 

                            

                            if (workSheet.Cells[f, 6].Value != null)
                            {
                                if (workSheet.Cells[f, 6].Value.ToString().Trim() != "")
                                    s.cups20 = workSheet.Cells[f, 6].Value.ToString();
                            }
                                                        

                            s.referencia = workSheet.Cells[f, 7].Value.ToString(); 
                            lista.Add(s);

                        }

                    }
                }
                fs = null;
                excelPackage = null;
                firstOnly = true;

                id = 0;
                for (int i = 0; i < lista.Count; i++)
                {
                    id++;
                    if (firstOnly)
                    {
                        sb.Append("replace into mes12_previsiones (factoring, nif, cliente,");
                        sb.Append(" cupsree, referencia, created_by, created_date) values ");
                        firstOnly = false;
                    }
                    sb.Append("('").Append(lista[i].factoring).Append("',");
                    sb.Append("'").Append(lista[i].nif).Append("',");
                    sb.Append("'").Append(lista[i].nombre_cliente).Append("',");                        

                    if (lista[i].cups20 != null)
                        sb.Append("'").Append(lista[i].cups20).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append(lista[i].referencia).Append("',");

                    sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                    if (id == 250)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        id = 0;
                    }

                }

                if (id > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    id = 0;
                }

            }
            catch (Exception e)
            {
                MessageBox.Show("Error en la línea " + linea + " --> " + e.Message,
                  "Error en la importación de adjudicaciones",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        private bool EstructuraCorrecta(string cabecera)
        {
            if (cabecera.ToUpper().Trim() == "ENTIDADLNEMPRESA TITULARNIFCLIENTECCOUNIPSCUPSREEREFERENCIASECCONTROL")
                return true;
            else
                return false;
        }

        private Dictionary<string, string> CargaDatosExcelConReferencias()
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            Dictionary<string, string> d = new Dictionary<string, string>();


            strSql = "SELECT e.CFACTURA, e.REFERENCIA FROM mes12_excel e"
                + " WHERE e.REFERENCIA IS NOT NULL";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                d.Add(r["CFACTURA"].ToString(), r["REFERENCIA"].ToString());

            }
            db.CloseConnection();


            return d;

        }

        private Dictionary<string, string> CargaDatosExcelConNIFAdjudicaciones()
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            Dictionary<string, string> d = new Dictionary<string, string>();


            strSql = "SELECT e.CFACTURA, e.NIF_EN_MES13 FROM mes12_excel e"
                + " WHERE e.NIF_EN_MES13 IS NOT NULL";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                d.Add(r["CFACTURA"].ToString(), r["NIF_EN_MES13"].ToString());

            }
            db.CloseConnection();


            return d;

        }
        private string GetReferencia(string cfactura)
        {
            
            string o;
            if (dic_referencias.TryGetValue(cfactura, out o))
                return o;
            else
                return null;
        }

        private string GetReferenciaNIF(string cfactura)
        {

            string o;
            if (dic_referencias_nif.TryGetValue(cfactura, out o))
                return o;
            else
                return null;
        }

        private bool EstructuraCorrectaPrevision(string cabecera)
        {
            if (cabecera.ToUpper().Trim() == "LNEMPRESA TITULARNIFCLIENTECCOUNIPSCUPSREEREFERENCIA")
                return true;
            else
                return false;
        }
        
        private void BorradoTabla_Mes12_Excel()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "REPLACE INTO mes12_excel_hist"
                + " SELECT * FROM mes12_excel";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "DELETE FROM mes12_excel";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        public double total_registros_individuales()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            int total_registros = 0;

            strSql = "select count(*) as total from mes12_adjudicaciones where referencia like '%NR%'";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();

            while (r.Read())
            {
                total_registros = Convert.ToInt32(r["total"]);
            }

            db.CloseConnection();
            return total_registros;

        }

        public double total_registros_agrupadas()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            int total_registros = 0;

            strSql = "select count(*) as total from mes12_adjudicaciones where referencia like '%AG%'";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();

            while (r.Read())
            {
                total_registros = Convert.ToInt32(r["total"]);
            }

            db.CloseConnection();
            return total_registros;

        }

        

    }



   
}
