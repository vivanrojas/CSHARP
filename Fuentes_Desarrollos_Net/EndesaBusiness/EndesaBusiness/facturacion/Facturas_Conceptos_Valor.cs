using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{

    // Esta clase genera un informe sobre facturas con conceptos concretos e importes
    // concretos. Busca también si el concepto se ha corregido en una refacturacion
    // posterior.
    public class Facturas_Conceptos_Valor
    {

        utilidades.Param p;
        logs.Log ficheroLog;
                
        
        public Facturas_Conceptos_Valor(bool lanza_ya)
        {

            DateTime fd = new DateTime();
            List<int> lista_conceptos = new List<int>();
            List<string> lista_nifs = new List<string>();
            List<string> lista_tipos_factura = new List<string>();

            utilidades.Fechas f = new utilidades.Fechas();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Facturacion_MT_sin_EATR");

            
            ficheroLog.Add("==============");
            ficheroLog.Add("==============");
            ficheroLog.Add("INICIO PROCESO");
            ficheroLog.Add("==============");
            ficheroLog.Add("==============");
            ficheroLog.Add("");
            ficheroLog.Add("");

                

            p = new utilidades.Param("fact_mt_eatr_param", MySQLDB.Esquemas.FAC);
            fd = new DateTime(2021,01,01);

            ficheroLog.Add("Buscando facturas desde el " + fd.ToString("dd/MM/yyyy"));

            ficheroLog.Add("Conceptos de facturación a tratar " + p.GetValue("conceptos", DateTime.Now, DateTime.Now));
            ficheroLog.Add("Nifs a excluir de la facturación: " + p.GetValue("omitir_nifs", DateTime.Now, DateTime.Now));

            string[] conceptos = p.GetValue("conceptos", DateTime.Now, DateTime.Now).Split(';');
            string[] nifs = p.GetValue("omitir_nifs", DateTime.Now, DateTime.Now).Split(';');
            string[] tipos_factura = p.GetValue("tipo_facturas", DateTime.Now, DateTime.Now).Split(';');

            for (int i = 0; i < conceptos.Length; i++)
                lista_conceptos.Add(Convert.ToInt32(conceptos[i].Trim()));

            for (int i = 0; i < nifs.Length; i++)
                lista_nifs.Add(nifs[i].Trim());

            for (int i = 0; i < tipos_factura.Length; i++)
                lista_tipos_factura.Add(tipos_factura[i].Trim());

            Facturas_MT_con_EATR_0(fd, lista_conceptos, lista_nifs, lista_tipos_factura);
            List<EndesaEntity.facturacion.Factura> lista_facturas_con_EATR_0 = Facturas();


            Facturas_MT_sin_EATR(fd, lista_conceptos, lista_nifs, lista_tipos_factura);
            List<EndesaEntity.facturacion.Factura> lista_facturas_sin_EATR = Facturas();

            Informe(lista_facturas_con_EATR_0, lista_facturas_sin_EATR);                

            ficheroLog.Add("==============");
            ficheroLog.Add("==============");
            ficheroLog.Add("F I N  PROCESO");
            ficheroLog.Add("==============");
            ficheroLog.Add("==============");
            ficheroLog.Add("");
            ficheroLog.Add("");
                





        }


        private void Facturas_MT_con_EATR_0(DateTime fd, List<int> lista_conceptos, List<string> lista_nif, List<string> lista_tipos_factura)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            bool firstOnly = true;

            try
            {

                strSql = "DELETE FROM fact_mt_eatr_tmp";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "REPLACE into fact_mt_eatr_tmp"
                    + " SELECT"
                    + " fe.descripcion as EMPRESA, f.CNIFDNIC, f.DAPERSOC,"
                    + " f.CREFEREN, f.SECFACTU, tp.descripcion as TFACTURA,"
                    + " f.TESTFACT, f.CFACTURA, f.CFACTREC,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " f.CUPSREE, f.IFACTURA, f.VCUOVAFA,"
                    + " t.TCONFAC, c.DESCRIPCION_CORTA, c.DESCRIPCION,"
                    + " t.ICONFAC"
                    + " FROM fo f"
                    + " inner join fact.fo_empresas fe on"
                    + " fe.empresa_id = f.ID_ENTORNO"
                    + " INNER JOIN fact.fo_p_tipos_factura tp ON"
                    + " tp.id_tipo_factura = f.TFACTURA"
                    + " INNER JOIN fact.fo_tcon t ON"
                    + " t.CREFEREN = f.CREFEREN AND"
                    + " t.SECFACTU = f.SECFACTU AND"
                    + " t.TESTFACT = f.TESTFACT"
                    + " INNER JOIN fact.fo_cf c ON"
                    + " c.CONC_SCE = t.TCONFAC"
                    + " WHERE"
                    + " fe.descripcion in ('MT-Portugal')";


                firstOnly = true;
                for (int i = 0; i < lista_tipos_factura.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += " AND tp.descripcion IN ("
                            + "'" + lista_tipos_factura[i] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista_tipos_factura[i] + "'";

                }
                strSql += ")";

                strSql += " AND f.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " AND f.TIPONEGOCIO = 'L'";


                firstOnly = true;
                for (int i = 0; i < lista_conceptos.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += " AND t.TCONFAC IN ("
                            + lista_conceptos[i];
                        firstOnly = false;
                    }
                    else
                        strSql += "," + lista_conceptos[i];
                   
                }

                strSql += ")";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                firstOnly = true;
                strSql = "DELETE FROM fact_mt_eatr_tmp WHERE CNIFDNIC IN (";

                for (int i = 0; i < lista_nif.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += "'" + lista_nif[i] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista_nif[i] + "'";

                }

                strSql += ")";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                strSql = "delete from fact_mt_eatr_tmp where TESTFACT in ('A', 'S')";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "DELETE FROM fact_mt_eatr";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "REPLACE INTO fact_mt_eatr"
                    + " SELECT * FROM fact_mt_eatr_tmp f"
                    + " ORDER BY f.CEMPTITU, f.CNIFDNIC, f.CREFEREN," 
                    + " f.CUPSREE, f.FFACTDES, f.SECFACTU";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "DELETE FROM fact_mt_eatr"
                    + " WHERE ICONFAC <> 0";                    
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Facturas_MT_con_EATR_0: " + e.Message);
            }
        }


        private void Facturas_MT_sin_EATR(DateTime fd, List<int> lista_conceptos, List<string> lista_nif, List<string> lista_tipos_factura)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            bool firstOnly = true;

            try
            {
                strSql = "DELETE FROM fact_mt_eatr_tmp";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "REPLACE into fact_mt_eatr_tmp"
                    + " SELECT"
                    + " fe.descripcion as EMPRESA, f.CNIFDNIC, f.DAPERSOC,"
                    + " f.CREFEREN, f.SECFACTU, tp.descripcion as TFACTURA,"
                    + " f.TESTFACT, f.CFACTURA, f.CFACTREC,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " f.CUPSREE, f.IFACTURA, f.VCUOVAFA,"
                    + " t.TCONFAC, c.DESCRIPCION_CORTA, c.DESCRIPCION,"
                    + " t.ICONFAC"
                    + " FROM fo f"
                    + " inner join fact.fo_empresas fe on"
                    + " fe.empresa_id = f.ID_ENTORNO"
                    + " INNER JOIN fact.fo_p_tipos_factura tp ON"
                    + " tp.id_tipo_factura = f.TFACTURA"
                    + " left outer JOIN fact.fo_tcon t ON"
                    + " t.CREFEREN = f.CREFEREN AND"
                    + " t.SECFACTU = f.SECFACTU AND"
                    + " t.TESTFACT = f.TESTFACT"

                //    for (int i = 0; i < lista_conceptos.Count; i++)
                //    {
                //        if (firstOnly)
                //        {
                //            strSql += " AND t.TCONFAC IN ("
                //                + lista_conceptos[i];
                //            firstOnly = false;
                //        }
                //        else
                //            strSql += "," + lista_conceptos[i];
                //    }

                //strSql += ")"



                    + " left outer JOIN fact.fo_cf c ON"
                    + " c.CONC_SCE = t.TCONFAC"
                    + " WHERE"
                    + " fe.descripcion in ('MT-Portugal')";

                firstOnly = true;
                for (int i = 0; i < lista_tipos_factura.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += " AND tp.descripcion IN ("
                            + "'" + lista_tipos_factura[i] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista_tipos_factura[i] + "'";

                }
                strSql += ")";


                strSql += " and f.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " AND f.TIPONEGOCIO = 'L'"
                    + " and f.VCUOVAFA <> 0"
                    + " and t.CREFEREN is null";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                firstOnly = true;
                strSql = "DELETE FROM fact_mt_eatr_tmp WHERE CNIFDNIC IN (";

                for (int i = 0; i < lista_nif.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += "'" + lista_nif[i] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista_nif[i] + "'";

                }
                strSql += ")";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "delete from fact_mt_eatr_tmp where TESTFACT in ('A', 'S')";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "DELETE FROM fact_mt_eatr";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "REPLACE INTO fact_mt_eatr"
                    + " SELECT * FROM fact_mt_eatr_tmp f"
                    + " ORDER BY f.CEMPTITU, f.CNIFDNIC, f.CREFEREN,"
                    + " f.CUPSREE, f.FFACTDES, f.SECFACTU";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch(Exception e)
            {
                ficheroLog.AddError("Facturas_MT_sin_EATR: " + e.Message);
            }



        }


        private List<EndesaEntity.facturacion.Factura> Facturas()            
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            List<EndesaEntity.facturacion.Factura> d
                = new List<EndesaEntity.facturacion.Factura>();

            try
            {
                strSql = "SELECT"
                    + " CEMPTITU, CNIFDNIC, DAPERSOC,"
                    + " CREFEREN, SECFACTU, TFACTURA,"
                    + " TESTFACT, CFACTURA, CFACTREC,"
                    + " FFACTURA, FFACTDES, FFACTHAS,"
                    + " CUPSREE, IFACTURA, VCUOVAFA,"
                    + " TCONFAC, DESCRIPCION_CORTA, DESCRIPCION,"
                    + " ICONFAC"
                    + " FROM fact_mt_eatr";


                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.Factura c = new EndesaEntity.facturacion.Factura();
                    c.descripcion_empresa = r["CEMPTITU"].ToString();
                    c.cnifdnic = r["CNIFDNIC"].ToString();
                    c.dapersoc = r["DAPERSOC"].ToString();
                    c.creferen = Convert.ToInt64(r["CREFEREN"]);
                    c.secfactu = Convert.ToInt32(r["SECFACTU"]);
                    c.tfactura_descripcion = r["TFACTURA"].ToString();
                    c.testfact = r["TESTFACT"].ToString();

                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();
                    if (r["CFACTREC"] != System.DBNull.Value)
                        c.cfactrec = r["CFACTREC"].ToString();

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FFACTURA"]);
                    if (r["FFACTDES"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);
                    if (r["FFACTHAS"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cupsree = r["CUPSREE"].ToString();

                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"]);

                    if (r["VCUOVAFA"] != System.DBNull.Value)
                        c.vcuovafa = Convert.ToInt32(r["VCUOVAFA"]);

                    if (r["TCONFAC"] != System.DBNull.Value)
                        c.tconfac = Convert.ToInt32(r["TCONFAC"]);
                    if (r["DESCRIPCION"] != System.DBNull.Value)
                        c.tconfac_descripcion_larga = r["DESCRIPCION"].ToString();

                    if (r["DESCRIPCION_CORTA"] != System.DBNull.Value)
                        c.tconfac_descripcion_corta = r["DESCRIPCION_CORTA"].ToString();

                    if (r["ICONFAC"] != System.DBNull.Value)
                        c.iconfac = Convert.ToDouble(r["ICONFAC"]);

                    d.Add(c);
                    

                }
                db.CloseConnection();
                return d;
            }catch(Exception e)
            {
                ficheroLog.AddError("Facturas: " + e.Message);
                return null;
            }

        }


        private Dictionary<string, EndesaEntity.facturacion.Factura> FacturasRectificativas(
            DateTime fecha_periodo_desde,
            DateTime fecha_periodo_hasta, int concepto_sce, double importe_concepto)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.facturacion.Factura> d
                = new Dictionary<string, EndesaEntity.facturacion.Factura>();

            try
            {
                strSql = "SELECT"
                    + " f.CEMPTITU, f.CNIFDNIC, f.DAPERSOC,"
                    + " f.CREFEREN, f.SECFACTU, f.TFACTURA,"
                    + " f.TESTFACT, f.CFACTURA, f.CFACTREC,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " f.CUPSREE, f.IFACTURA,"
                    + " t.TCONFAC, c.DESCRIPCION_CORTA, c.DESCRIPCION,"
                    + " t.ICONFAC"
                    + " FROM fo f"
                    + " INNER JOIN fact.fo_tcon t on"
                    + " t.CREFEREN = f.CREFEREN and"
                    + " t.SECFACTU = f.SECFACTU and"
                    + " t.TESTFACT = f.TESTFACT"
                    + " INNER JOIN fact.fo_cf c on"
                    + " c.CONC_SCE = t.TCONFAC"
                    + " WHERE"
                    + " (f.FFACTDES >= '" + fecha_periodo_desde.ToString("yyyy-MM-dd") + "') and"
                    + " f.TESTFACT IN ('Y') AND"
                    + " (f.CFACTREC is not NULL AND f.CFACTREC <> '') and"
                    + " t.TCONFAC = " + concepto_sce + " AND"
                    + " t.ICONFAC <> " + importe_concepto.ToString().Replace(",", ".");

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.Factura c = new EndesaEntity.facturacion.Factura();
                    c.cemptitu = Convert.ToInt32(r["CEMPTITU"]);
                    c.cnifdnic = r["CNIFDNIC"].ToString();
                    c.creferen = Convert.ToInt64(r["CREFEREN"]);
                    c.secfactu = Convert.ToInt32(r["SECFACTU"]);
                    c.tfactura = Convert.ToInt32(r["TFACTURA"]);
                    c.testfact = r["TESTFACT"].ToString();

                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();
                    if (r["CFACTREC"] != System.DBNull.Value)
                        c.cfactrec = r["CFACTREC"].ToString();

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FFACTURA"]);
                    if (r["FFACTDES"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);
                    if (r["FFACTHAS"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cupsree = r["CUPSREE"].ToString();

                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"]);

                    EndesaEntity.facturacion.Factura_Detalle cc = new EndesaEntity.facturacion.Factura_Detalle();
                    cc.tconfac = Convert.ToInt32(r["TCONFAC"]);
                    cc.iconfac = Convert.ToDouble(r["ICONFAC"]);
                    cc.tconfac_desc_corta = r["DESCRIPCION_CORTA"].ToString();
                    cc.tcofac_desc_larga = r["DESCRIPCION"].ToString();

                    c.lista_conceptos.Add(cc);

                    EndesaEntity.facturacion.Factura o;
                    if (!d.TryGetValue(c.cfactrec, out o))
                        d.Add(c.cfactrec, c);
                    else
                        Console.WriteLine("Existe");

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }

        }


        


        private void Informe(List<EndesaEntity.facturacion.Factura> lf_eatr_0,
            List<EndesaEntity.facturacion.Factura> lf_sin_eatr)
        {

            int f = 0;
            int c = 0;

            string[] listaArchivos = Directory.GetFiles(p.GetValue("ubicacion_salida_informe", DateTime.Now, DateTime.Now),
                  p.GetValue("prefijo_informe_excel", DateTime.Now, DateTime.Now)  + "*.xlsx");


            if (listaArchivos.Length > 0)
            {
                for (int i = 0; i < listaArchivos.Length; i++)
                {
                    FileInfo fichero = new FileInfo(listaArchivos[i]);
                    fichero.Delete();
                }
                    
            }

           FileInfo fileInfo = new FileInfo(p.GetValue("ubicacion_salida_informe",DateTime.Now, DateTime.Now)
                + p.GetValue("prefijo_informe_excel", DateTime.Now, DateTime.Now)
                + DateTime.Now.ToString("yyyy_MM_dd_HHmmss") + ".xlsx");


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Facturas con EATR 0");

            var headerCells = workSheet.Cells[1, 1, 1, 19];
            var headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;
            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "CEMPTITU"; c++;
            workSheet.Cells[f, c].Value = "CNIFDNIC"; c++;
            workSheet.Cells[f, c].Value = "DAPERSOC"; c++;
            workSheet.Cells[f, c].Value = "CREFEREN"; c++;
            workSheet.Cells[f, c].Value = "SECFACTU"; c++;
            workSheet.Cells[f, c].Value = "TFACTURA"; c++;
            workSheet.Cells[f, c].Value = "TESTFACT"; c++;
            workSheet.Cells[f, c].Value = "CFACTURA"; c++;
            workSheet.Cells[f, c].Value = "CFACTREC"; c++;
            workSheet.Cells[f, c].Value = "FFACTURA"; c++;
            workSheet.Cells[f, c].Value = "FFACTDES"; c++;
            workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
            workSheet.Cells[f, c].Value = "CUPSREE"; c++;
            workSheet.Cells[f, c].Value = "IFACTURA"; c++;
            workSheet.Cells[f, c].Value = "VCUOVAFA"; c++;
            workSheet.Cells[f, c].Value = "TCONFAC"; c++;
            workSheet.Cells[f, c].Value = "DESCRIPCIÓN CORTA"; c++;
            workSheet.Cells[f, c].Value = "DESCRIPCIÓN LARGA"; c++;
            workSheet.Cells[f, c].Value = "IMPORTE CONCEPTO"; 
            


            for (int i = 1; i <= c; i++)
            {
                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }

            foreach (EndesaEntity.facturacion.Factura p in lf_eatr_0)
            {
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = p.descripcion_empresa; c++;
                workSheet.Cells[f, c].Value = p.cnifdnic; c++;
                workSheet.Cells[f, c].Value = p.dapersoc; c++;
                workSheet.Cells[f, c].Value = p.creferen; c++;
                workSheet.Cells[f, c].Value = p.secfactu; c++;
                workSheet.Cells[f, c].Value = p.tfactura_descripcion; c++;
                workSheet.Cells[f, c].Value = p.testfact; c++;
                workSheet.Cells[f, c].Value = p.cfactura; c++;
                workSheet.Cells[f, c].Value = p.cfactrec; c++;

                if (p.ffactura > DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = p.ffactura;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (p.ffactdes > DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = p.ffactdes;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (p.ffacthas > DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = p.ffacthas;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;                
                workSheet.Cells[f, c].Value = p.cupsree; c++;
                workSheet.Cells[f, c].Value = p.ifactura;
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = p.vcuovafa;
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.tconfac; c++;
                workSheet.Cells[f, c].Value = p.tconfac_descripcion_corta; c++;
                workSheet.Cells[f, c].Value = p.tconfac_descripcion_larga; c++;
                workSheet.Cells[f, c].Value = p.iconfac;
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

            }

            var allCells = workSheet.Cells[1, 1, f, 19];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:S1"].AutoFilter = true;
            allCells.AutoFitColumns();


            workSheet = excelPackage.Workbook.Worksheets.Add("Facturas sin EATR");
            headerCells = workSheet.Cells[1, 1, 1, 15];
            headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;
            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "CEMPTITU"; c++;
            workSheet.Cells[f, c].Value = "CNIFDNIC"; c++;
            workSheet.Cells[f, c].Value = "DAPERSOC"; c++;
            workSheet.Cells[f, c].Value = "CREFEREN"; c++;
            workSheet.Cells[f, c].Value = "SECFACTU"; c++;
            workSheet.Cells[f, c].Value = "TFACTURA"; c++;
            workSheet.Cells[f, c].Value = "TESTFACT"; c++;
            workSheet.Cells[f, c].Value = "CFACTURA"; c++;
            workSheet.Cells[f, c].Value = "CFACTREC"; c++;
            workSheet.Cells[f, c].Value = "FFACTURA"; c++;
            workSheet.Cells[f, c].Value = "FFACTDES"; c++;
            workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
            workSheet.Cells[f, c].Value = "CUPSREE"; c++;
            workSheet.Cells[f, c].Value = "IFACTURA"; c++;
            workSheet.Cells[f, c].Value = "VCUOVAFA"; 



            for (int i = 1; i <= c; i++)
            {
                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }

            foreach (EndesaEntity.facturacion.Factura p in lf_sin_eatr)
            {
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = p.descripcion_empresa; c++;
                workSheet.Cells[f, c].Value = p.cnifdnic; c++;
                workSheet.Cells[f, c].Value = p.dapersoc; c++;
                workSheet.Cells[f, c].Value = p.creferen; c++;
                workSheet.Cells[f, c].Value = p.secfactu; c++;
                workSheet.Cells[f, c].Value = p.tfactura_descripcion; c++;
                workSheet.Cells[f, c].Value = p.testfact; c++;
                workSheet.Cells[f, c].Value = p.cfactura; c++;
                workSheet.Cells[f, c].Value = p.cfactrec; c++;

                if (p.ffactura > DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = p.ffactura;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (p.ffactdes > DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = p.ffactdes;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (p.ffacthas > DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = p.ffacthas;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;
                workSheet.Cells[f, c].Value = p.cupsree; c++;
                workSheet.Cells[f, c].Value = p.ifactura;
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = p.vcuovafa;
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
            }
               

            allCells = workSheet.Cells[1, 1, f, 15];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:O1"].AutoFilter = true;
            allCells.AutoFitColumns();

            excelPackage.Save();

            GeneraCorreo(fileInfo.FullName, lf_eatr_0, lf_sin_eatr);
        }


        private void GeneraCorreo(string archivo, List<EndesaEntity.facturacion.Factura> lf_eatr_0,
            List<EndesaEntity.facturacion.Factura> lf_sin_eatr)
        {
            string body = "";
            string subject = "";
            string from = "";
            string to = "";
            string cc = null;
            string attachment = null;


            try
            {
                from = p.GetValue("remitente", DateTime.Now, DateTime.Now);
                to = p.GetValue("destinatarios_mail_para", DateTime.Now, DateTime.Now);
                subject = p.GetValue("asunto_mail", DateTime.Now, DateTime.Now);

                if(lf_eatr_0.Count > 0 || lf_sin_eatr.Count > 0)
                {
                    attachment = archivo;
                    body = (DateTime.Now.Hour > 14 ? "Buenas tardes." : "Buenos días.")
                        + System.Environment.NewLine
                        + "     Se han detectado " + lf_eatr_0.Count + " facturas de MT con EATR = 0 y "
                        + System.Environment.NewLine
                        + "se han detectado " + lf_sin_eatr.Count + " facturas de MT sin EATR."
                        + System.Environment.NewLine
                        + "Un saludo.";

                }
                else
                {
                    body = DateTime.Now.Hour > 14 ? "Buenas tardes." : "Buenos días."
                        + System.Environment.NewLine
                        + "     No se han detectado facturas de MT con EATR = 0 ni facturas sin EATR."
                        + System.Environment.NewLine
                        + "Un saludo.";
                }



                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                mes.SendMail(from, to, cc, subject, body, attachment);

            }
            catch(Exception e)
            {
                ficheroLog.AddError("GeneraCorreo: " + e.Message);
            }
        }

    }
}
