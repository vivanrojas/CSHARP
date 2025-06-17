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
using System.Windows.Forms;

namespace EndesaBusiness.facturacion 
{
    public class Facturacion : EndesaEntity.facturacion.Factura
    {
        Dictionary<string, List<EndesaEntity.facturacion.Factura>> dic;
        public Facturacion()
        {
            
        }

        public void Busca_Facturas(DateTime fd_factura, DateTime fh_factura, List<string> lista_cups20,
            List<string> lista_segmentos, List<string> lista_empresas, List<string> lista_tfacturas)
        {
            dic = DIC_Facturas(fd_factura, fh_factura, lista_cups20, lista_segmentos, lista_empresas, lista_tfacturas);
        }

        private List<EndesaEntity.facturacion.Factura> Facturas(
            DateTime fd_factura, DateTime fh_factura,
            List<string> lista_segmentos, List<string> lista_empresas, List<string> lista_tfacturas)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            bool firstOnly = true;

            List<EndesaEntity.facturacion.Factura> l =
                new List<EndesaEntity.facturacion.Factura>();
            try
            {
                strSql = "SELECT e.descripcion AS empresa, f.CNIFDNIC, f.DAPERSOC, f.CUPSREE,"
                     + " f.CREFEREN, f.SECFACTU, f.TESTFACT,"
                     + " f.IVA, f.IIMPUES2, f.IIMPUES3, f.CCOUNIPS,"
                     + " f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS, f.VCUOVAFA,"
                     + " f.IFACTURA, f.TESTFACT, tf.descripcion as TFACTURA"
                     + " FROM fo f"
                     + " INNER JOIN fo_empresas e ON"
                     + " e.empresa_id = f.ID_ENTORNO"
                     + " INNER JOIN fo_p_tipos_factura tf ON"
                     + " tf.id_tipo_factura = f.TFACTURA"
                     + " WHERE";

                for (int i = 0; i < lista_empresas.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += " e.descripcion IN ('" + lista_empresas[i] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista_empresas[i] + "'";
                }

                firstOnly = true;
                strSql += ") AND (f.FFACTURA >= '" + fd_factura.ToString("yyyy-MM-dd") + "' AND"
                     + " f.FFACTURA <= '" + fh_factura.ToString("yyyy-MM-dd") + "')";                    

                for (int i = 0; i < lista_segmentos.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += " AND f.TIPONEGOCIO IN ('" + lista_segmentos[i] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista_segmentos[i] + "'";

                }
                if(lista_segmentos.Count > 0)
                    strSql += ")";

                firstOnly = true;
                for (int i = 0; i < lista_tfacturas.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += " AND tf.descripcion IN ('" + lista_tfacturas[i] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista_tfacturas[i] + "'";

                }

                if (lista_tfacturas.Count > 0)
                    strSql += ")";

                Console.WriteLine(strSql);
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.Factura c = new EndesaEntity.facturacion.Factura();

                    c.descripcion_empresa = r["empresa"].ToString();
                    c.cnifdnic = r["CNIFDNIC"].ToString();
                    c.dapersoc = r["DAPERSOC"].ToString();

                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        c.ccounips = r["CCOUNIPS"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cupsree = r["CUPSREE"].ToString();

                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                    if (r["FFACTDES"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                    if (r["FFACTHAS"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);

                    if (r["VCUOVAFA"] != System.DBNull.Value)
                        c.vcuovafa = Convert.ToInt32(r["VCUOVAFA"]);

                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"]);

                    if (r["IVA"] != System.DBNull.Value)
                        c.iva = Convert.ToDouble(r["IVA"]);

                    if (r["IIMPUES2"] != System.DBNull.Value)
                        c.iimpues2 = Convert.ToDouble(r["IIMPUES2"]);

                    if (r["IIMPUES3"] != System.DBNull.Value)
                        c.iimpues3 = Convert.ToDouble(r["IIMPUES3"]);

                    if (r["TESTFACT"] != System.DBNull.Value)
                        c.testfact = r["TESTFACT"].ToString();

                    if (r["TFACTURA"] != System.DBNull.Value)
                        c.tfactura_desc = r["TFACTURA"].ToString();

                    if (r["CREFEREN"] != System.DBNull.Value)
                        c.creferen = Convert.ToInt64(r["CREFEREN"]);

                    if (r["SECFACTU"] != System.DBNull.Value)
                        c.secfactu = Convert.ToInt32(r["SECFACTU"]);

                    if (r["TESTFACT"] != System.DBNull.Value)
                        c.testfact = r["TESTFACT"].ToString();


                    l.Add(c);


                }
                db.CloseConnection();
                return l;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                   "Estructura de columnas Excel incorrecta",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);

                return null;
            }
        }


        private Dictionary<string, List<EndesaEntity.facturacion.Factura>> DIC_Facturas(
           DateTime fd_factura, DateTime fh_factura, List<string> lista_cups20,
           List<string> lista_segmentos, List<string> lista_empresas, List<string> lista_tfacturas)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            bool firstOnly = true;

            Dictionary<string, List<EndesaEntity.facturacion.Factura>> d
                = new Dictionary<string, List<EndesaEntity.facturacion.Factura>>();


            try
            {
                strSql = "SELECT e.descripcion AS empresa, f.CNIFDNIC, f.DAPERSOC, f.CUPSREE,"
                     + " f.CREFEREN, f.SECFACTU, f.TESTFACT,"
                     + " f.IVA, f.IIMPUES2, f.IIMPUES3, f.CCOUNIPS,"
                     + " f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS, f.VCUOVAFA,"
                     + " f.IFACTURA, f.TESTFACT, tf.descripcion as TFACTURA"
                     + " FROM fo f"
                     + " INNER JOIN fo_empresas e ON"
                     + " e.empresa_id = f.ID_ENTORNO"
                     + " INNER JOIN fo_p_tipos_factura tf ON"
                     + " tf.id_tipo_factura = f.TFACTURA"
                     + " WHERE";


                for (int i = 0; i < lista_cups20.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += " f.CUPSREE IN ('" + lista_cups20[i] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista_cups20[i] + "'";
                }

                firstOnly = true;
                for (int i = 0; i < lista_empresas.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += " e.descripcion IN ('" + lista_empresas[i] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista_empresas[i] + "'";
                }

                firstOnly = true;
                strSql += ") AND (f.FFACTURA >= '" + fd_factura.ToString("yyyy-MM-dd") + "' AND"
                     + " f.FFACTURA <= '" + fh_factura.ToString("yyyy-MM-dd") + "')";

                for (int i = 0; i < lista_segmentos.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += " AND f.TIPONEGOCIO IN ('" + lista_segmentos[i] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista_segmentos[i] + "'";

                }
                if (lista_segmentos.Count > 0)
                    strSql += ")";

                firstOnly = true;
                for (int i = 0; i < lista_tfacturas.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += " AND tf.descripcion IN ('" + lista_tfacturas[i] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista_tfacturas[i] + "'";

                }

                if (lista_tfacturas.Count > 0)
                    strSql += ")";

                strSql += " order by f.CUPSREE, f.FFACTDES desc";

                Console.WriteLine(strSql);
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.Factura c = new EndesaEntity.facturacion.Factura();

                    c.descripcion_empresa = r["empresa"].ToString();
                    c.cnifdnic = r["CNIFDNIC"].ToString();
                    c.dapersoc = r["DAPERSOC"].ToString();

                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        c.ccounips = r["CCOUNIPS"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cupsree = r["CUPSREE"].ToString();

                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                    if (r["FFACTDES"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                    if (r["FFACTHAS"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);

                    if (r["VCUOVAFA"] != System.DBNull.Value)
                        c.vcuovafa = Convert.ToInt32(r["VCUOVAFA"]);

                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"]);

                    if (r["IVA"] != System.DBNull.Value)
                        c.iva = Convert.ToDouble(r["IVA"]);

                    if (r["IIMPUES2"] != System.DBNull.Value)
                        c.iimpues2 = Convert.ToDouble(r["IIMPUES2"]);

                    if (r["IIMPUES3"] != System.DBNull.Value)
                        c.iimpues3 = Convert.ToDouble(r["IIMPUES3"]);

                    if (r["TESTFACT"] != System.DBNull.Value)
                        c.testfact = r["TESTFACT"].ToString();

                    if (r["TFACTURA"] != System.DBNull.Value)
                        c.tfactura_desc = r["TFACTURA"].ToString();

                    if (r["CREFEREN"] != System.DBNull.Value)
                        c.creferen = Convert.ToInt64(r["CREFEREN"]);

                    if (r["SECFACTU"] != System.DBNull.Value)
                        c.secfactu = Convert.ToInt32(r["SECFACTU"]);

                    if (r["TESTFACT"] != System.DBNull.Value)
                        c.testfact = r["TESTFACT"].ToString();


                    if(c.cupsree != "" && c.cupsree.Length >= 20)
                    {
                        List<EndesaEntity.facturacion.Factura> o;
                        if (!d.TryGetValue(c.cupsree.Substring(0,20), out o))
                        {
                            o = new List<EndesaEntity.facturacion.Factura>();
                            o.Add(c);
                            d.Add(c.cupsree.Substring(0, 20), o);
                        }
                        else
                            o.Add(c);
                    }

                    


                }
                db.CloseConnection();
                return d;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                   "Estructura de columnas Excel incorrecta",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);

                return null;
            }
        }


        public void InformeExcel_FacturasGas_PT_Real_Estimada(DateTime fd_factura, DateTime fh_factura)
        {

            int c = 1;
            int f = 1;
            SaveFileDialog save;
            List<string> lista_segmentos = new List<string>();
            List<string> lista_empresas = new List<string>();
            List<string> lista_tfacturas = new List<string>();
            EndesaBusiness.sigame.SIGAME s;

            try
            {

                s = new sigame.SIGAME(fd_factura, fh_factura, true, "PT");

                lista_segmentos.Add("G");
                lista_empresas.Add("BTN-Portugal");
                lista_empresas.Add("BTE-Portugal");
                lista_empresas.Add("MT-Portugal");

                

                Cursor.Current = Cursors.WaitCursor;
                List<EndesaEntity.facturacion.Factura> lista =
                           Facturas(fd_factura, fh_factura, lista_segmentos, lista_empresas, lista_tfacturas);
                
                Cursor.Current = Cursors.Default;

                if (lista.Count() > 0)
                {


                    save = new SaveFileDialog();
                    save.Title = "Ubicación del informe";
                    save.AddExtension = true;
                    save.DefaultExt = "xlsx";
                    DialogResult result = save.ShowDialog();
                    if (result == DialogResult.OK)
                    {

                        FileInfo fileInfo = new FileInfo(save.FileName);
                        if (fileInfo.Exists)
                            fileInfo.Delete();

                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                        ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                        var workSheet = excelPackage.Workbook.Worksheets.Add("FACTURAS");
                        var headerCells = workSheet.Cells[1, 1, 1, 47];
                        var headerFont = headerCells.Style.Font;

                        workSheet.View.FreezePanes(2, 1);

                        headerFont.Bold = true;
                        workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                        workSheet.Cells[f, c].Value = "CIF"; c++;
                        workSheet.Cells[f, c].Value = "CLIENTE"; c++;
                        workSheet.Cells[f, c].Value = "CUI"; c++;
                        workSheet.Cells[f, c].Value = "COD. FACTURA"; c++;
                        workSheet.Cells[f, c].Value = "FECHA FACTURA"; c++;
                        workSheet.Cells[f, c].Value = "FECHA DESDE "; c++;
                        workSheet.Cells[f, c].Value = "FECHA HASTA"; c++;
                        workSheet.Cells[f, c].Value = "CONSUMO"; c++;
                        workSheet.Cells[f, c].Value = "IMPORTE TOTAL"; c++;
                        workSheet.Cells[f, c].Value = "TIPO FACTURA"; c++;
                        workSheet.Cells[f, c].Value = "FUENTE"; c++;
                        workSheet.Cells[f, c].Value = "MEDIDA"; c++;

                        foreach(EndesaEntity.facturacion.Factura p in lista)
                        {
                            f++;
                            c = 1;
                            workSheet.Cells[f, c].Value = p.descripcion_empresa; c++;
                            workSheet.Cells[f, c].Value = p.cnifdnic; c++;
                            workSheet.Cells[f, c].Value = p.dapersoc; c++;
                            workSheet.Cells[f, c].Value = p.cupsree; c++;
                            workSheet.Cells[f, c].Value = p.cfactura; c++;

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

                            workSheet.Cells[f, c].Value = p.vcuovafa;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            c++;

                            workSheet.Cells[f, c].Value = p.ifactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;
                            workSheet.Cells[f, c].Value = p.testfact; c++;

                            if(p.cfactura != null)
                                workSheet.Cells[f, c].Value = s.Fuente(p.cfactura);
                            c++;

                            if (p.cfactura != null)
                                workSheet.Cells[f, c].Value = s.Medida(p.cfactura);
                            c++;
                        }

                        var allCells = workSheet.Cells[1, 1, f, 12];
                        workSheet.Cells["A1:M1"].AutoFilter = true;


                        headerFont.Bold = true;
                        allCells.AutoFitColumns();
                        excelPackage.Save();

                        MessageBox.Show("Informe terminado.",
                               "Informe Excel",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Information);

                        DialogResult result3 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                        if (result3 == DialogResult.Yes)
                            System.Diagnostics.Process.Start(save.FileName);

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                               "Informe Excel",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
            }
        }

        public void InformeExcel_FacturasGas_ES_Real_Estimada(DateTime fd_factura, DateTime fh_factura)
        {

            int c = 1;
            int f = 1;
            SaveFileDialog save;
            List<string> lista_segmentos = new List<string>();
            List<string> lista_empresas = new List<string>();
            List<string> lista_tfacturas = new List<string>();
            EndesaBusiness.sigame.SIGAME s;

            try
            {

                s = new sigame.SIGAME(fd_factura, fh_factura, true, "ES");

                lista_segmentos.Add("G");
                lista_empresas.Add("MT-España");
                lista_empresas.Add("EEXXI");
                
                Cursor.Current = Cursors.WaitCursor;
                List<EndesaEntity.facturacion.Factura> lista =
                           Facturas(fd_factura, fh_factura, lista_segmentos, lista_empresas, lista_tfacturas);

                Cursor.Current = Cursors.Default;

                if (lista.Count() > 0)
                {


                    save = new SaveFileDialog();
                    save.Title = "Ubicación del informe";
                    save.AddExtension = true;
                    save.DefaultExt = "xlsx";
                    DialogResult result = save.ShowDialog();
                    if (result == DialogResult.OK)
                    {

                        FileInfo fileInfo = new FileInfo(save.FileName);
                        if (fileInfo.Exists)
                            fileInfo.Delete();

                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                        ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                        var workSheet = excelPackage.Workbook.Worksheets.Add("FACTURAS");
                        var headerCells = workSheet.Cells[1, 1, 1, 47];
                        var headerFont = headerCells.Style.Font;

                        workSheet.View.FreezePanes(2, 1);

                        headerFont.Bold = true;
                        workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                        workSheet.Cells[f, c].Value = "CIF"; c++;
                        workSheet.Cells[f, c].Value = "CLIENTE"; c++;
                        workSheet.Cells[f, c].Value = "CUI"; c++;
                        workSheet.Cells[f, c].Value = "COD. FACTURA"; c++;
                        workSheet.Cells[f, c].Value = "FECHA FACTURA"; c++;
                        workSheet.Cells[f, c].Value = "FECHA DESDE "; c++;
                        workSheet.Cells[f, c].Value = "FECHA HASTA"; c++;
                        workSheet.Cells[f, c].Value = "CONSUMO"; c++;
                        workSheet.Cells[f, c].Value = "IMPORTE TOTAL"; c++;
                        workSheet.Cells[f, c].Value = "TIPO FACTURA"; c++;
                        workSheet.Cells[f, c].Value = "FUENTE"; c++;
                        workSheet.Cells[f, c].Value = "MEDIDA"; c++;

                        foreach (EndesaEntity.facturacion.Factura p in lista)
                        {
                            f++;
                            c = 1;
                            workSheet.Cells[f, c].Value = p.descripcion_empresa; c++;
                            workSheet.Cells[f, c].Value = p.cnifdnic; c++;
                            workSheet.Cells[f, c].Value = p.dapersoc; c++;
                            workSheet.Cells[f, c].Value = p.cupsree; c++;
                            workSheet.Cells[f, c].Value = p.cfactura; c++;

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

                            workSheet.Cells[f, c].Value = p.vcuovafa;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            c++;

                            workSheet.Cells[f, c].Value = p.ifactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;
                            workSheet.Cells[f, c].Value = p.testfact; c++;

                            if (p.cfactura != null)
                                workSheet.Cells[f, c].Value = s.Fuente(p.cfactura);
                            c++;

                            if (p.cfactura != null)
                                workSheet.Cells[f, c].Value = s.Medida(p.cfactura);
                            c++;
                        }

                        var allCells = workSheet.Cells[1, 1, f, 12];
                        workSheet.Cells["A1:M1"].AutoFilter = true;


                        headerFont.Bold = true;
                        allCells.AutoFitColumns();
                        excelPackage.Save();

                        MessageBox.Show("Informe terminado.",
                               "Informe Excel",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Information);

                        DialogResult result3 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                        if (result3 == DialogResult.Yes)
                            System.Diagnostics.Process.Start(save.FileName);

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                               "Informe Excel",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
            }
        }

        public void InformeExcel_FacturasElectricidad_BTN_Real_Estimada(DateTime fd_factura, DateTime fh_factura)
        {

            int c = 1;
            int f = 1;
            SaveFileDialog save;
            List<string> lista_segmentos = new List<string>();
            List<string> lista_empresas = new List<string>();
            List<string> lista_tfacturas = new List<string>();
            Dictionary<string, string> dic_cups13 = new Dictionary<string, string>();
            DateTime min_fecha = new DateTime();
            DateTime max_fecha = new DateTime();
            bool firstOnly = true;

            try
            {
                                

                lista_segmentos.Add("L");

                lista_empresas.Add("BTN-Portugal");

                lista_tfacturas.Add("CICLO");
                lista_tfacturas.Add("MANUAL");
                lista_tfacturas.Add("PSEUDOFACTURA");
                lista_tfacturas.Add("MODIFICACION CONTRATO");
                lista_tfacturas.Add("BAJA CONTRATO");
                lista_tfacturas.Add("FACT. ESPECIAL DE CONTRATACION"); 

                Cursor.Current = Cursors.WaitCursor;
                List<EndesaEntity.facturacion.Factura> lista =
                           Facturas(fd_factura, fh_factura, lista_segmentos, lista_empresas, lista_tfacturas);

                
                foreach(EndesaEntity.facturacion.Factura p in lista)
                {
                    if(p.ccounips != null)
                    {
                        if (firstOnly)
                        {
                            min_fecha = p.ffactdes;
                            max_fecha = p.ffacthas;
                            firstOnly = false;
                        }

                        min_fecha = (p.ffactdes < min_fecha ? p.ffactdes : min_fecha);
                        max_fecha = (p.ffacthas > max_fecha ? p.ffacthas : max_fecha);

                        string o;


                        if (!dic_cups13.TryGetValue(p.ccounips, out o))
                            dic_cups13.Add(p.ccounips, p.ccounips);
                    }
                    
                }


                LecturasElectricidadBTN lecturas = 
                    new LecturasElectricidadBTN(dic_cups13.Select(z => z.Key).ToList(),min_fecha,max_fecha);


                Cursor.Current = Cursors.Default;

                if (lista.Count() > 0)
                {


                    save = new SaveFileDialog();
                    save.Title = "Ubicación del informe";
                    save.AddExtension = true;
                    save.DefaultExt = "xlsx";
                    DialogResult result = save.ShowDialog();
                    if (result == DialogResult.OK)
                    {

                        FileInfo fileInfo = new FileInfo(save.FileName);
                        if (fileInfo.Exists)
                            fileInfo.Delete();

                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                        ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                        var workSheet = excelPackage.Workbook.Worksheets.Add("FACTURAS");
                        var headerCells = workSheet.Cells[1, 1, 1, 47];
                        var headerFont = headerCells.Style.Font;

                        workSheet.View.FreezePanes(2, 1);

                        headerFont.Bold = true;
                        workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                        workSheet.Cells[f, c].Value = "CIF"; c++;
                        workSheet.Cells[f, c].Value = "CLIENTE"; c++;
                        workSheet.Cells[f, c].Value = "CUPSREE"; c++;
                        workSheet.Cells[f, c].Value = "COD. FACTURA"; c++;
                        workSheet.Cells[f, c].Value = "FECHA FACTURA"; c++;
                        workSheet.Cells[f, c].Value = "FECHA DESDE "; c++;
                        workSheet.Cells[f, c].Value = "FECHA HASTA"; c++;
                        workSheet.Cells[f, c].Value = "CONSUMO"; c++;
                        workSheet.Cells[f, c].Value = "IVA"; c++;
                        workSheet.Cells[f, c].Value = "IIMPUES2"; c++;
                        workSheet.Cells[f, c].Value = "IMP_CAV"; c++;
                        workSheet.Cells[f, c].Value = "IMPORTE TOTAL"; c++;
                        workSheet.Cells[f, c].Value = "CREFEREN"; c++;
                        workSheet.Cells[f, c].Value = "SECFACTU"; c++;
                        workSheet.Cells[f, c].Value = "TESTFACT"; c++;
                        workSheet.Cells[f, c].Value = "TIPO FACTURA"; c++;
                        workSheet.Cells[f, c].Value = "CUPS CORTO"; c++;
                        workSheet.Cells[f, c].Value = "TFUENTE SCE"; c++;
                        workSheet.Cells[f, c].Value = "DESCRIPCION FUENTE SCE"; c++;
                        workSheet.Cells[f, c].Value = "COMENTARIO"; c++;
                        workSheet.Cells[f, c].Value = "FUENTE"; c++;

                        foreach (EndesaEntity.facturacion.Factura p in lista)
                        {
                            f++;
                            c = 1;
                            workSheet.Cells[f, c].Value = p.descripcion_empresa; c++;
                            workSheet.Cells[f, c].Value = p.cnifdnic; c++;
                            workSheet.Cells[f, c].Value = p.dapersoc; c++;
                            workSheet.Cells[f, c].Value = p.cupsree; c++;
                            workSheet.Cells[f, c].Value = p.cfactura; c++;

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

                            workSheet.Cells[f, c].Value = p.vcuovafa;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            c++;

                            if (p.iva != 0)
                            {
                                workSheet.Cells[f, c].Value = p.iva;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                
                            }
                            c++;

                            if (p.iimpues2 != 0)
                            {
                                workSheet.Cells[f, c].Value = p.iimpues2;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";

                            }
                            c++;

                            if (p.iimpues3 != 0)
                            {
                                workSheet.Cells[f, c].Value = p.iimpues3;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";

                            }
                            c++;

                            workSheet.Cells[f, c].Value = p.creferen; c++;
                            workSheet.Cells[f, c].Value = p.secfactu; c++;
                            workSheet.Cells[f, c].Value = p.testfact; c++;

                            workSheet.Cells[f, c].Value = p.ifactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;
                            workSheet.Cells[f, c].Value = p.tfactura_desc; c++;

                            workSheet.Cells[f, c].Value = p.ccounips; c++;

                            EndesaEntity.medida.LecturasElectricidad lectura =
                                lecturas.GetLectura(p.ccounips, p.ffacthas);

                            if(lectura != null)
                            {
                                
                                workSheet.Cells[f, c].Value = lectura.tfuente; c++;
                                workSheet.Cells[f, c].Value = lectura.descripcion_fuente; c++;
                                c++; // Comentario
                                workSheet.Cells[f, c].Value = lectura.fuente;

                            }
                            else
                            {
                                c++; c++;c++;
                            }

                        }

                        var allCells = workSheet.Cells[1, 1, f, c];
                        workSheet.Cells["A1:V1"].AutoFilter = true;


                        headerFont.Bold = true;
                        allCells.AutoFitColumns();
                        excelPackage.Save();

                        MessageBox.Show("Informe terminado.",
                               "Informe Excel",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Information);

                        DialogResult result3 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                        if (result3 == DialogResult.Yes)
                            System.Diagnostics.Process.Start(save.FileName);

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                               "Informe Excel",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
            }
        }

        public void GetUltimaFacturaEmitida(string cups20) 
        {
            List<EndesaEntity.facturacion.Factura> o;
            if (dic.TryGetValue(cups20, out o))
            {
                this.ffactura = o[0].ffactura;
                this.ffactdes = o[0].ffactdes;
                this.ffacthas = o[0].ffacthas;
            }


        }

    }
}
