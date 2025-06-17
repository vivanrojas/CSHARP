using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{
    public class SIGAME_Importe_Redes
    {

        public List<EndesaEntity.facturacion.InformeGasPortugal> lista_facturas { get; set; }
        List<EndesaEntity.facturacion.Conceptos> lista_conceptos = new List<EndesaEntity.facturacion.Conceptos>();
        List<EndesaEntity.facturacion.ConceptosEspeciales> lce = new List<EndesaEntity.facturacion.ConceptosEspeciales>();
        utilidades.Param p = new utilidades.Param("fact_p_exchange_param", MySQLDB.Esquemas.FAC);
        // Desconto DL 84-D/2022
        Dictionary<string, EndesaEntity.facturacion.Conceptos> dic_dl84 = new Dictionary<string, EndesaEntity.facturacion.Conceptos>();

        public SIGAME_Importe_Redes()
        {
            lista_facturas = new List<EndesaEntity.facturacion.InformeGasPortugal>();
        }

        public void CargardgvCUPS(bool usarFechaFactura, DateTime fd, DateTime fh, string cupsree)
        {

            string strSql = "";
            SQLServer db;
            SqlCommand command;
            SqlDataReader reader;
            
            DateTime min_fd = fd;
            DateTime max_fh = fh;
            bool existe_dif = false;

            try
            {

                lista_facturas.Clear();

                #region query
                strSql = "SELECT ps.ID_PS, ps.CD_CUPS, c.DE_NOMBRE_CLIENTE, c.CD_CIF, "
                + " fr.FH_INI_FACTURACION, fr.FH_FIN_FACTURACION, fr.CD_TESTFACT,"
                + " fr.CD_NFACTURA_REALES_PS, fr.FH_FACTURA,"
                + " fr.NM_IMPORTE_NETO, fr.NM_IMPORTE_IVA, fr.NM_IMPORTE_IVA2, fr.NM_IMPORTE_IVA3,"
                + " fr.NM_IMPORTE_BRUTO, fr.CD_MONEDA,"
                + " fr.NM_QFACTURADA, fr.NM_CONSUMO, fr.NM_TOS_IMPORTE, null as dif, null as Consumo_m3, null as Consumo_kwh,"
                + " fr.NM_IEH_IMPORTE, cont.LG_PF,cont.LG_INDEX_POOL, cont.LG_INDEX_BRENT"
                + " FROM (T_SGM_G_PS ps INNER JOIN (T_SGM_M_FACTURAS_REALES_PS fr INNER JOIN T_SGM_G_CONTRATOS_PS cont ON"
                + " fr.ID_CTO_PS = cont.ID_CTO_PS) ON ps.ID_PS = cont.ID_PS) INNER JOIN T_SGM_M_CLIENTES c ON"
                + " ps.ID_CLIENTE = c.ID_CLIENTE"
                + " WHERE ps.CD_CUPS Like 'PT%' AND (fr.CD_TESTFACT in ('N','S','Y') OR fr.CD_TESTFACT is null) AND"
                + " fr.CD_NFACTURA_REALES_PS is not null AND";

                if (usarFechaFactura)
                {
                    strSql = strSql + " (fr.FH_FACTURA >= '" + String.Format("{0:yyyyMMdd}", fd) + "' and"
                        + "  fr.FH_FACTURA <= '" + String.Format("{0:yyyyMMdd}", fh) + "')";
                }
                else
                {
                    strSql = strSql + " (fr.FH_INI_FACTURACION >= '" + String.Format("{0:yyyyMMdd}", fd) + "' and"
                        + " fr.FH_FIN_FACTURACION <= '" + String.Format("{0:yyyyMMdd}", fh) + "')";
                }

                if (cupsree != null)
                {
                    strSql = strSql + " and ps.CD_CUPS = '" + cupsree + "'";
                }

                strSql = strSql + " order by fr.FH_INI_FACTURACION, ps.CD_CUPS ";
                #endregion

                Cursor.Current = Cursors.WaitCursor;

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);

                SqlDataAdapter da = new SqlDataAdapter(command);
                //DataTable dt = new DataTable();
                //da.Fill(dt);
                //dgvCups.DataSource = dt;

                reader = command.ExecuteReader();


                while (reader.Read())
                {
                    EndesaEntity.facturacion.InformeGasPortugal c = new EndesaEntity.facturacion.InformeGasPortugal();

                    c.id_ps = Convert.ToInt32(reader["ID_PS"]);
                    c.cups20 = reader["CD_CUPS"].ToString();
                    c.dapersoc = reader["DE_NOMBRE_CLIENTE"].ToString();
                    c.nif = reader["CD_CIF"].ToString();

                    c.ffactdes = Convert.ToDateTime(reader["FH_INI_FACTURACION"]);
                    c.ffacthas = Convert.ToDateTime(reader["FH_FIN_FACTURACION"]);

                   


                    if (min_fd > c.ffactdes)
                        min_fd = c.ffactdes;

                    if (max_fh < c.ffacthas)
                        max_fh = c.ffacthas;

                    if (reader["CD_NFACTURA_REALES_PS"] != System.DBNull.Value)
                        c.cfactura = reader["CD_NFACTURA_REALES_PS"].ToString();

                    if (reader["FH_FACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(reader["FH_FACTURA"]);

                    c.importe_neto = Convert.ToDouble(reader["NM_IMPORTE_NETO"]);
                    c.importe_iva = Convert.ToDouble(reader["NM_IMPORTE_IVA"]);
                    c.importe_bruto = Convert.ToDouble(reader["NM_IMPORTE_BRUTO"]);

                    if (reader["NM_IMPORTE_IVA2"] != System.DBNull.Value)
                        c.importe_iva_reducido = Convert.ToDouble(reader["NM_IMPORTE_IVA2"]);

                    c.moneda = reader["CD_MONEDA"].ToString();

                    if (reader["NM_QFACTURADA"] != System.DBNull.Value)
                        c.qfacturada = Convert.ToDouble(reader["NM_QFACTURADA"]);

                    if (reader["NM_CONSUMO"] != System.DBNull.Value)
                        c.consumo = Convert.ToDouble(reader["NM_CONSUMO"]);

                    if (reader["NM_TOS_IMPORTE"] != System.DBNull.Value)
                        c.tos_importe = Convert.ToDouble(reader["NM_TOS_IMPORTE"]);

                    if (reader["NM_IEH_IMPORTE"] != System.DBNull.Value)
                        c.ieh_importe = Convert.ToDouble(reader["NM_IEH_IMPORTE"]);
                    if (reader["Consumo_kwh"] != System.DBNull.Value)
                        c.consumo_kwh = Convert.ToDouble(reader["Consumo_kwh"]);

                    if (reader["Consumo_kwh"] != System.DBNull.Value)
                        c.consumo_kwh = Convert.ToDouble(reader["Consumo_kwh"]);


                    if (reader["Dif"] != System.DBNull.Value)
                        c.dif = Convert.ToDouble(reader["Dif"]);

                    if (c.nif != "A82434697")
                        c.import_acceso_redes = this.ImporteAccesoRedes(c.cfactura);
                    else
                        c.import_acceso_redes = this.ImporteAccesoRedesEspecial(c.ffactdes, c.ffacthas);

                    if (reader["LG_PF"] != System.DBNull.Value)
                        c.lg_pf = Convert.ToBoolean(reader["LG_PF"]);

                    if (reader["LG_INDEX_POOL"] != System.DBNull.Value)
                        c.lg_pool = Convert.ToBoolean(reader["LG_INDEX_POOL"]);

                    if (reader["LG_INDEX_BRENT"] != System.DBNull.Value)
                        c.lg_brent = Convert.ToBoolean(reader["LG_INDEX_BRENT"]);

                    EndesaEntity.facturacion.Conceptos o;
                    if(dic_dl84.TryGetValue(c.cfactura, out o))
                    {
                        c.importe_dto = o.importe;
                        c.kw_dto = o.kw;
                    }


                    lista_facturas.Add(c);

                }
                db.CloseConnection();

                if (lista_facturas.Count == 0)
                {
                    MessageBox.Show("La consulta que ha realizado no devuelve ningún resultado.",
                    "La consulta no devuelte datos.",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                }
                else
                {
                    // Buscamos los consumos segun las fechas 
                    sigame.SIGAME sig = new sigame.SIGAME(min_fd, max_fh);
                    for (int j = 0; j < lista_facturas.Count(); j++)
                    {

                        sig.GetConsumo(lista_facturas[j].id_ps, Convert.ToInt32(lista_facturas[j].ffactdes.ToString("yyyyMM")));
                        lista_facturas[j].consumo_m3 = sig.consumos.consumo_ruta + sig.consumos.consumo_tm;
                        lista_facturas[j].consumo_kwh = sig.consumos.kwh_ruta + sig.consumos.kwh_tm;
                        if (lista_facturas[j].consumo < 0)
                            lista_facturas[j].dif = 0;
                        else
                        {
                            lista_facturas[j].dif = (lista_facturas[j].consumo_kwh - lista_facturas[j].consumo);
                            if (!existe_dif)
                                existe_dif = lista_facturas[j].dif != 0;
                        }

                    }

                    if (existe_dif)
                        MessageBox.Show("La consulta que ha realizado contiene diferencias." +
                            System.Environment.NewLine + "Por favor, revise el listado buscando por la columna DIF.",
                   "Diferencias en Consumos.",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Information);
                }


                
                Cursor.Current = Cursors.Default;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la importación de CUPS",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

        }

        

        public string UltimaImportacion()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;            
            string ult = null;
            DateTime f = new DateTime();

            try
            {
                strSql = "select r.time_period, r.value from fact_p_exchange_rates r order by r.time_period desc limit 1";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    f = Convert.ToDateTime(reader["time_period"]);
                    ult = f.ToString("dd/MM/yyyy") + " --> 1 EUR = " + reader["value"].ToString() + " USD";
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return ult;
        }

        

        public void ExportExcelOpenOffice(string rutaFichero)
        {
            int f = 0;
            int c = 0;

            string nif = "";
            List<EndesaEntity.facturacion.InformeGasPortugal> inf = new List<EndesaEntity.facturacion.InformeGasPortugal>();
            try
            {
                FileInfo file = new FileInfo(rutaFichero);
                if (file.Exists)
                    file.Delete();

                nif = p.GetValue("nif_elecgas", DateTime.Now,DateTime.Now);

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(file);                
                var workSheet = excelPackage.Workbook.Worksheets.Add("FACTURADO");


                #region Resumen

                f = 1;
                c = 1;

                var headerCells = workSheet.Cells[1, 1, 1, 25];
                var headerFont = headerCells.Style.Font;
                headerFont.Bold = true;

                workSheet.Cells[f, c].Value = "CD_CUPS"; c++;
                workSheet.Cells[f, c].Value = "DE_NOMBRE_CLIENTE"; c++;
                workSheet.Cells[f, c].Value = "CD_CIF"; c++;
                workSheet.Cells[f, c].Value = "FH_INI_FACTURACION"; c++;
                workSheet.Cells[f, c].Value = "FH_FIN_FACTURACION"; c++;
                workSheet.Cells[f, c].Value = "CD_NFACTURA_REALES_PS"; c++;
                workSheet.Cells[f, c].Value = "FH_FACTURA"; c++;
                workSheet.Cells[f, c].Value = "NM_IMPORTE_NETO"; c++;
                workSheet.Cells[f, c].Value = "NM_IMPORTE_IVA"; c++;
                workSheet.Cells[f, c].Value = "NM_IMPORTE_IVA_REDUCIDO"; c++;
                workSheet.Cells[f, c].Value = "NM_IMPORTE_BRUTO"; c++;
                workSheet.Cells[f, c].Value = "NM_QFACTURADA"; c++;
                workSheet.Cells[f, c].Value = "NM_CONSUMO"; c++;
                workSheet.Cells[f, c].Value = "Consumo m3 cargado"; c++;
                workSheet.Cells[f, c].Value = "NM_TOS_IMPORTE"; c++;
                workSheet.Cells[f, c].Value = "NM_IEH_IMPORTE"; c++;
                workSheet.Cells[f, c].Value = "Consumo kWh cargado"; c++;
                workSheet.Cells[f, c].Value = "Dif"; c++;
                workSheet.Cells[f, c].Value = "Importe Acceso a Redes sin IVA"; c++;
                workSheet.Cells[f, c].Value = "Descuento DL 84-D/2022"; c++;
                workSheet.Cells[f, c].Value = "Consumo Descuento"; c++;
                workSheet.Cells[f, c].Value = "Precio Fijo"; c++;
                workSheet.Cells[f, c].Value = "Precio indexado Pool"; c++;
                workSheet.Cells[f, c].Value = "Precio indexado Brent"; c++;

                inf = lista_facturas.FindAll(z => z.nif != nif);

                for (int i = 0; i < inf.Count(); i++)
                {
                    c = 1;
                    f++;
                    workSheet.Cells[f, c].Value = inf[i].cups20; c++;
                    workSheet.Cells[f, c].Value = inf[i].dapersoc; c++;
                    workSheet.Cells[f, c].Value = inf[i].nif; c++;
                    workSheet.Cells[f, c].Value = inf[i].ffactdes; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = inf[i].ffacthas; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;                    
                    workSheet.Cells[f, c].Value = inf[i].cfactura; c++;

                    if(inf[i].ffactura > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = inf[i].ffactura; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;
                    workSheet.Cells[f, c].Value = inf[i].importe_neto; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = inf[i].importe_iva; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    if (inf[i].importe_iva_reducido != 0)
                    {
                        workSheet.Cells[f, c].Value = inf[i].importe_iva_reducido; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    }
                    c++;

                    workSheet.Cells[f, c].Value = inf[i].importe_bruto; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = inf[i].qfacturada; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = inf[i].consumo; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = inf[i].consumo_m3; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = inf[i].tos_importe; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = inf[i].ieh_importe; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = inf[i].consumo_kwh; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = inf[i].dif; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = inf[i].import_acceso_redes; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = inf[i].importe_dto; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = inf[i].kw_dto; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    if (inf[i].lg_pf)
                        workSheet.Cells[f, c].Value = "S";
                    else
                        workSheet.Cells[f, c].Value = "N";

                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    if (inf[i].lg_pool)
                        workSheet.Cells[f, c].Value = "S";
                    else
                        workSheet.Cells[f, c].Value = "N";

                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    if (inf[i].lg_brent)
                        workSheet.Cells[f, c].Value = "S";
                    else
                        workSheet.Cells[f, c].Value = "N";

                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;


                    if (inf[i].dif != 0)
                        for (int z = 1; z < c; z++)
                        {
                            workSheet.Cells[f, z].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, z].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.SandyBrown);
                        }


                    if (inf[i].cfactura == "" || inf[i].cfactura == null || inf[i].ffactura == DateTime.MinValue)
                        for (int z = 1; z < c; z++)
                        {
                            workSheet.Cells[f, z].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, z].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PaleVioletRed);
                        }


                }

                var allCells = workSheet.Cells[1, 1, f, c];
                allCells.AutoFitColumns();
                workSheet.Cells["A1:W1"].AutoFilter = true;
                workSheet.View.FreezePanes(2, 1);

                #endregion


                #region ELECGAS €
                workSheet = excelPackage.Workbook.Worksheets.Add("ELECGAS €");

                f = 1;
                c = 1;

                headerCells = workSheet.Cells[1, 1, 1, 23];
                headerFont = headerCells.Style.Font;

                workSheet.Cells[f, c].Value = "CD_CUPS"; c++;
                workSheet.Cells[f, c].Value = "DE_NOMBRE_CLIENTE"; c++;
                workSheet.Cells[f, c].Value = "CD_CIF"; c++;
                workSheet.Cells[f, c].Value = "FH_INI_FACTURACION"; c++;
                workSheet.Cells[f, c].Value = "FH_FIN_FACTURACION"; c++;
                workSheet.Cells[f, c].Value = "CD_NFACTURA_REALES_PS"; c++;
                workSheet.Cells[f, c].Value = "FH_FACTURA"; c++;
                workSheet.Cells[f, c].Value = "NM_IMPORTE_NETO"; c++;
                workSheet.Cells[f, c].Value = "NM_IMPORTE_IVA"; c++;
                workSheet.Cells[f, c].Value = "NM_IMPORTE_BRUTO"; c++;
                workSheet.Cells[f, c].Value = "NM_QFACTURADA"; c++;
                workSheet.Cells[f, c].Value = "NM_CONSUMO"; c++;
                workSheet.Cells[f, c].Value = "Consumo m3 cargado"; c++;
                workSheet.Cells[f, c].Value = "Consumo kWh cargado"; c++;
                workSheet.Cells[f, c].Value = "Dif"; c++;
                workSheet.Cells[f, c].Value = "Importe Acceso a Redes sin IVA"; c++;
                workSheet.Cells[f, c].Value = "Precio Fijo"; c++;
                workSheet.Cells[f, c].Value = "Precio indexado Pool"; c++;
                workSheet.Cells[f, c].Value = "Precio indexado Brent"; c++;

                inf = lista_facturas.FindAll(z => z.nif == nif && z.moneda == "EUR");

                for (int i = 0; i < inf.Count(); i++)
                {
                    c = 1;
                    f++;
                    workSheet.Cells[f, c].Value = inf[i].cups20; c++;
                    workSheet.Cells[f, c].Value = inf[i].dapersoc; c++;
                    workSheet.Cells[f, c].Value = inf[i].nif; c++;
                    workSheet.Cells[f, c].Value = inf[i].ffactdes; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = inf[i].ffacthas; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = inf[i].cfactura; c++;
                    workSheet.Cells[f, c].Value = inf[i].ffactura; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = inf[i].importe_neto; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = inf[i].importe_iva; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = inf[i].importe_bruto; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = inf[i].qfacturada; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = inf[i].consumo; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = inf[i].consumo_m3; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = inf[i].consumo_kwh; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = inf[i].dif; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = inf[i].import_acceso_redes; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    if (inf[i].lg_pf)
                        workSheet.Cells[f, c].Value = "S";
                    else
                        workSheet.Cells[f, c].Value = "N";

                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    if (inf[i].lg_pool)
                        workSheet.Cells[f, c].Value = "S";
                    else
                        workSheet.Cells[f, c].Value = "N";

                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    if (inf[i].lg_brent)
                        workSheet.Cells[f, c].Value = "S";
                    else
                        workSheet.Cells[f, c].Value = "N";

                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    if (inf[i].dif != 0)
                        for (int z = 1; z < c; z++)
                        {
                            workSheet.Cells[f, z].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, z].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.SandyBrown);
                        }

                    if (inf[i].cfactura == "" || inf[i].cfactura == null || inf[i].ffactura == DateTime.MinValue)
                        for (int z = 1; z < c; z++)
                        {
                            workSheet.Cells[f, z].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, z].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PaleVioletRed);
                        }
                }

                headerFont.Bold = true;
                allCells = workSheet.Cells[1, 1, f, c];
                allCells.AutoFitColumns();
                workSheet.Cells["A1:S1"].AutoFilter = true;
                workSheet.View.FreezePanes(2, 1);

                #endregion

                #region ELECGAS $
                workSheet = excelPackage.Workbook.Worksheets.Add("ELECGAS $");

                f = 1;
                c = 1;

                headerCells = workSheet.Cells[1, 1, 1, 23];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;

                workSheet.Cells[f, c].Value = "CD_CUPS"; c++;
                workSheet.Cells[f, c].Value = "DE_NOMBRE_CLIENTE"; c++;
                workSheet.Cells[f, c].Value = "CD_CIF"; c++;
                workSheet.Cells[f, c].Value = "FH_INI_FACTURACION"; c++;
                workSheet.Cells[f, c].Value = "FH_FIN_FACTURACION"; c++;
                workSheet.Cells[f, c].Value = "CD_NFACTURA_REALES_PS"; c++;
                workSheet.Cells[f, c].Value = "FH_FACTURA"; c++;
                workSheet.Cells[f, c].Value = "NM_IMPORTE_NETO"; c++;
                workSheet.Cells[f, c].Value = "NM_IMPORTE_IVA"; c++;
                workSheet.Cells[f, c].Value = "NM_IMPORTE_BRUTO $"; c++;
                workSheet.Cells[f, c].Value = "NM_IMPORTE_BRUTO €"; c++;
                workSheet.Cells[f, c].Value = "NM_QFACTURADA"; c++;
                workSheet.Cells[f, c].Value = "NM_CONSUMO"; c++;
                workSheet.Cells[f, c].Value = "Consumo m3 cargado"; c++;
                workSheet.Cells[f, c].Value = "Consumo kWh cargado"; c++;
                workSheet.Cells[f, c].Value = "Dif"; c++;
                workSheet.Cells[f, c].Value = "Precio Fijo"; c++;
                workSheet.Cells[f, c].Value = "Precio indexado Pool"; c++;
                workSheet.Cells[f, c].Value = "Precio indexado Brent"; c++;

                inf = lista_facturas.FindAll(z => z.nif == nif && z.moneda == "DOL");

                for (int i = 0; i < inf.Count(); i++)
                {
                    c = 1;
                    f++;
                    workSheet.Cells[f, c].Value = inf[i].cups20; c++;
                    workSheet.Cells[f, c].Value = inf[i].dapersoc; c++;
                    workSheet.Cells[f, c].Value = inf[i].nif; c++;
                    workSheet.Cells[f, c].Value = inf[i].ffactdes; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = inf[i].ffacthas; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = inf[i].cfactura; c++;
                    workSheet.Cells[f, c].Value = inf[i].ffactura; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = inf[i].importe_neto; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = inf[i].importe_iva; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = inf[i].importe_bruto; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Math.Round(inf[i].importe_bruto / this.GetDolarValue(inf[i].ffactura), 2);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = inf[i].qfacturada; c++;
                    workSheet.Cells[f, c].Value = inf[i].consumo; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = inf[i].consumo_m3; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = inf[i].consumo_kwh; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = inf[i].dif; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    if (inf[i].lg_pf)
                        workSheet.Cells[f, c].Value = "S";
                    else
                        workSheet.Cells[f, c].Value = "N";

                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    if (inf[i].lg_pool)
                        workSheet.Cells[f, c].Value = "S";
                    else
                        workSheet.Cells[f, c].Value = "N";

                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    if (inf[i].lg_brent)
                        workSheet.Cells[f, c].Value = "S";
                    else
                        workSheet.Cells[f, c].Value = "N";

                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    if (inf[i].dif != 0)
                        for (int z = 1; z < c; z++)
                        {
                            workSheet.Cells[f, z].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, z].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.SandyBrown);
                        }

                    if (inf[i].cfactura == "" || inf[i].cfactura == null || inf[i].ffactura == DateTime.MinValue)
                        for (int z = 1; z < c; z++)
                        {
                            workSheet.Cells[f, z].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, z].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PaleVioletRed);
                        }

                }


                #endregion


                headerFont.Bold = true;
                allCells = workSheet.Cells[1, 1, f, c];
                workSheet.Cells["A1:S1"].AutoFilter = true;
                allCells.AutoFitColumns();
                excelPackage.Save();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Exportar a Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CargaImporteConceptosEspeciales(bool usarFechaFactura, DateTime fd, DateTime fh)
        {
            string strSql = "";
            SQLServer db;
            SqlCommand command;
            SqlDataReader reader;
            string cupsree_especial = null;

            try
            {

                cupsree_especial = p.GetValue("cupsree_elecgas", DateTime.Now, DateTime.Now);

                strSql = "select T_SGM_M_CTO_BONIFICAC.* from T_SGM_M_CTO_BONIFICAC,T_SGM_P_TIPO_BONIFICACION,"
                + " T_SGM_G_PS,T_SGM_G_CONTRATOS_PS"
                + " where cd_cups = '" + cupsree_especial + "'"
                + " and T_SGM_G_PS.id_ps = T_SGM_G_CONTRATOS_PS.id_ps"
                + " and T_SGM_M_CTO_BONIFICAC.id_cto_ps = T_SGM_G_CONTRATOS_PS.id_cto_ps"
                + " and T_SGM_P_TIPO_BONIFICACION.id_tipo_bonificacion = T_SGM_M_CTO_BONIFICAC.id_tipo_bonificacion";
               

                lce.Clear();
                Cursor.Current = Cursors.WaitCursor;
                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    EndesaEntity.facturacion.ConceptosEspeciales c = new EndesaEntity.facturacion.ConceptosEspeciales();
                    if (reader["FH_INICIO"] != System.DBNull.Value)
                        c.fd = Convert.ToDateTime(reader["FH_INICIO"]);

                    if (reader["FH_FIN"] != System.DBNull.Value)
                        c.fh = Convert.ToDateTime(reader["FH_FIN"]);

                    if (reader["NM_DTO_EUROS"] != System.DBNull.Value)
                        c.descuento = Convert.ToDouble(reader["NM_DTO_EUROS"]);
                    lce.Add(c);
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la importación de Conceptos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        public void CargaImporte_DTO(bool usarFechaFactura, DateTime fd, DateTime fh)
        {
            string strSql = "";
            SQLServer db;
            SqlCommand command;
            SqlDataReader reader;
            

            try
            {
                strSql = "SELECT fr.CD_NFACTURA_REALES_PS, r20.CodConcepto, r20.Descripcion,"                    
                    + " CASE"
                    + " WHEN RIGHT(r20.Importe, 1) = '1' THEN"
                        + " CONVERT(DECIMAL, LEFT(r20.Importe, LEN(r20.Importe) - 1)) / 100"
                    + " ELSE (CONVERT(DECIMAL, LEFT(r20.Importe, LEN(r20.Importe) - 1)) / 100) * -1"
                    + " END as ImporteConcepto, NM_CANTIDAD"                    
                    + " FROM T_SGM_M_FACTURAS_REALES_PS fr INNER JOIN Int_SCE_R_20 r20 ON"
                    + " r20.NumRefFactura = fr.CD_CREFAEXT"                    
                    + " WHERE";

                if (usarFechaFactura)
                {
                    strSql = strSql + " (fr.FH_FACTURA >= '" + String.Format("{0:yyyyMMdd}", fd) + "' and"
                        + "  fr.FH_FACTURA <= '" + String.Format("{0:yyyyMMdd}", fh) + "') and";
                }
                else
                {
                    strSql = strSql + " ( fr.FH_INI_FACTURACION >= '" + String.Format("{0:yyyyMMdd}", fd) + "' and"
                        + " fr.FH_FIN_FACTURACION <= '" + String.Format("{0:yyyyMMdd}", fh) + "') and";
                }

                strSql = strSql + " r20.CodConcepto = 'GB51'";                

                //lista_conceptos.Clear();
                Cursor.Current = Cursors.WaitCursor;
                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    EndesaEntity.facturacion.Conceptos c = new EndesaEntity.facturacion.Conceptos();
                    c.cfactura = reader["CD_NFACTURA_REALES_PS"].ToString();
                    c.importe = Convert.ToDouble(reader["ImporteConcepto"]);
                    c.kw = Convert.ToDouble(reader["NM_CANTIDAD"]);

                    EndesaEntity.facturacion.Conceptos o;
                    if (!dic_dl84.TryGetValue(c.cfactura, out o))
                        dic_dl84.Add(c.cfactura, c);
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la importación de Conceptos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

        }

        public void CargaImportesConceptos(bool usarFechaFactura, DateTime fd, DateTime fh, string cupsree)
        {
            string strSql = "";
            SQLServer db;
            SqlCommand command;
            SqlDataReader reader;


            try
            {
                strSql = "SELECT r10.FH_ULT_ACTUALIZACION, r10.AliasPS, r10.CIFNIF,"
                    + "r10.NumRefFactura, r10.CUPS22, fr.NM_IMPORTE_BRUTO,"
                    + "r10.NumContrato, r20.NumRefFactura, fr.NM_IMPORTE_NETO,"
                    + "r30.Importe, fr.ID_CTO_PS, fr.NM_CONSUMO,"
                    + "r30.ConsumoTotal, fr.CD_NFACTURA_REALES_PS, fr.FH_FACTURA,"
                    + "fr.CD_ESTADO, fr.TX_TIPO_FACTURA_NUEVO, fr.ID_FACTURA, fr.FH_INI_FACTURACION,"
                    + "fr.FH_FIN_FACTURACION, fr.NM_TOS_IMPORTE, r20.CodConcepto, r20.Descripcion,"
                    + "CASE"
                    + " WHEN RIGHT(r20.Importe, 1) = '1' THEN"
                        + " CONVERT(DECIMAL, LEFT(r20.Importe, LEN(r20.Importe) - 1)) / 100"
                    + " ELSE(CONVERT(DECIMAL, LEFT(r20.Importe, LEN(r20.Importe) - 1)) / 100) * -1"
                    + " END as ImporteConcepto,"
                    + " r10.FH_ULT_ACTUALIZACION"
                    + " FROM(T_SGM_M_FACTURAS_REALES_PS fr INNER JOIN(Int_SCE_R_10 r10 INNER JOIN"
                    + " Int_SCE_R_20 r20 ON r10.NumRefFactura = r20.NumRefFactura) ON fr.CD_CREFAEXT = r10.NumRefFactura)"
                    + " INNER JOIN Int_SCE_R_30 r30 ON r20.NumRefFactura = r30.NumRefFactura"
                    + " WHERE";

                if (usarFechaFactura)
                {
                    strSql = strSql + " (fr.FH_FACTURA >= '" + String.Format("{0:yyyyMMdd}", fd) + "' and"
                        + "  fr.FH_FACTURA <= '" + String.Format("{0:yyyyMMdd}", fh) + "') and";
                }
                else
                {
                    strSql = strSql + " ( fr.FH_INI_FACTURACION >= '" + String.Format("{0:yyyyMMdd}", fd) + "' and"
                        + " fr.FH_FIN_FACTURACION <= '" + String.Format("{0:yyyyMMdd}", fh) + "') and";
                }

                strSql = strSql + " r10.CUPS22 like 'PT%' AND"
                + " (r20.Descripcion Like '%T. Fixo%' Or r20.Descripcion Like 'Taxa fixa%' Or"
                + " r20.Descripcion Like 'Tº Energía (Tarifa)%' Or"
                + " r20.Descripcion Like 'Tº da Energia (Taxa)%')";

                if (cupsree != null)
                {
                    strSql = strSql + " and substring(r10.CUPS22,1,20) = '" + cupsree + "'";
                }

                lista_conceptos.Clear();
                Cursor.Current = Cursors.WaitCursor;
                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    EndesaEntity.facturacion.Conceptos c = new EndesaEntity.facturacion.Conceptos();
                    c.cfactura = reader["CD_NFACTURA_REALES_PS"].ToString();
                    c.importe = Convert.ToDouble(reader["ImporteConcepto"]);
                    lista_conceptos.Add(c);
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la importación de Conceptos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        private double GetDolarValue(DateTime f)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;            
            double value = 1;

            try
            {
                strSql = "Select value from fact_p_exchange_rates where"
                   + " time_period = '" + f.ToString("yyyy-MM-dd") + "' and"
                   + " currency = 'USD'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    value = Convert.ToDouble(reader["value"]);
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return value;
        }

        private double ImporteAccesoRedesEspecial(DateTime fd, DateTime fh)
        {
            double importe = 0;

            List<EndesaEntity.facturacion.ConceptosEspeciales> c = new List<EndesaEntity.facturacion.ConceptosEspeciales>();

            c = lce.FindAll(z => z.fd == fd && z.fh == fh);

            for (int i = 0; i < c.Count(); i++)
            {
                importe += c[i].descuento * -1;
            }

            return importe;

        }

        private double ImporteAccesoRedes(string cfactura)
        {
            double importe = 0;
            List<EndesaEntity.facturacion.Conceptos> c = new List<EndesaEntity.facturacion.Conceptos>();

            c = lista_conceptos.FindAll(z => z.cfactura == cfactura);

            for (int i = 0; i < c.Count(); i++)
            {
                importe += c[i].importe;
            }

            return importe;
        }

    }
}
