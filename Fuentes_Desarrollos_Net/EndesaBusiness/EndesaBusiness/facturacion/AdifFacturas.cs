using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using MySqlX.XDevAPI.Common;
using OfficeOpenXml;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{
    public class AdifFacturas
    {
        System.Data.DataTable dt;
        EndesaBusiness.utilidades.Param p;
        public List<EndesaEntity.facturacion.Adif_Factura> lista_facturas { get; set; }
        
            
        public AdifFacturas()
        {
            lista_facturas = new List<EndesaEntity.facturacion.Adif_Factura>();
            p = new utilidades.Param("adif_param", MySQLDB.Esquemas.FAC);
        }

        public AdifFacturas(DateTime ffactdes, DateTime ffacthas, string cupsree, string lote, bool diferencias)
        {
            p = new utilidades.Param("adif_param", MySQLDB.Esquemas.FAC);
            lista_facturas = CargaCups(ffactdes,  ffacthas,  cupsree,  lote,  diferencias);
            
        }

        private List<EndesaEntity.facturacion.Adif_Factura> CargaCups(DateTime ffactdes, DateTime ffacthas, string cupsree, string lote, bool diferencias)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            Dictionary<string, EndesaEntity.facturacion.Factura> dic_facturas_endesa;
            Dictionary<string, EndesaEntity.facturacion.Adif_Factura> dic_facturas_adif;

            Adif_calculaImpuestos impuestos = new Adif_calculaImpuestos();
            Dictionary<string, EndesaEntity.facturacion.Adif_Factura> d =
                new Dictionary<string, EndesaEntity.facturacion.Adif_Factura>();

            List<EndesaEntity.facturacion.Adif_Factura> lista = new List<EndesaEntity.facturacion.Adif_Factura>();

            EndesaBusiness.adif.InventarioFunciones inventario;


            try
            {
                
              
                List<string> lista_lotes = new List<string>();
                if (lote != "")
                    lista_lotes.Add(lote);

                if(cupsree == "")
                    inventario = new adif.InventarioFunciones(ffactdes, ffacthas, null, lista_lotes);
                else
                    inventario = new adif.InventarioFunciones(ffactdes, ffacthas, cupsree, lista_lotes);

                if (cupsree == "")
                    dic_facturas_endesa = Facturas_BI(ffactdes, ffacthas, null);
                else
                    dic_facturas_endesa = Facturas_BI(ffactdes, ffacthas, cupsree);

                dic_facturas_adif = Facturas_ADIF(ffactdes, ffacthas);
                
                foreach (KeyValuePair<string, EndesaEntity.medida.AdifInventario>  p in inventario.dic_inventario)
                {
                    EndesaEntity.facturacion.Adif_Factura c =
                       new EndesaEntity.facturacion.Adif_Factura();

                    c.CUPSREE = p.Value.cups20;
                    c.LOTE = p.Value.lote;
                    c.medida_en_baja = p.Value.medida_en_baja;
                    c.devolucion_de_energia = p.Value.devolucion_de_energia;
                    c.cierres_energia = p.Value.cierres_energia;

                    EndesaEntity.facturacion.Adif_Factura o_adif;
                    if (dic_facturas_adif.TryGetValue(c.CUPSREE, out o_adif))
                    {
                        c.existe_factura_adif = true;
                        c.CONSUMO_ADIF = o_adif.CONSUMO_ADIF;
                        c.cnpr_adif = o_adif.cnpr_adif + o_adif.cpre_adif;
                    }

                    EndesaEntity.facturacion.Factura o_endesa;
                    if (dic_facturas_endesa.TryGetValue(c.CUPSREE, out o_endesa))
                    {

                        c.existe_factura_sce = true;
                        c.CONSUMO_SCE = o_endesa.vcuovafa;
                        c.cnpr_endesa = o_endesa.cnpr;


                        c.CREFEREN = o_endesa.creferen.ToString();
                        c.SECFACTU = o_endesa.secfactu;
                        c.FFACTURA = o_endesa.ffactura;
                        c.FFACTDES = o_endesa.ffactdes;
                        c.FFACTHAS = o_endesa.ffacthas;
                        c.testfact = o_endesa.testfact;
                        c.de_tfactura = o_endesa.tfactura_desc;

                        c.DIF_CONSUMO = c.CONSUMO_ADIF - c.CONSUMO_SCE;
                        c.DIF_TOTAL = (c.cnpr_adif + c.cpre_adif) - c.cnpr_endesa;

                    }
                    else
                    {
                        c.DIF_CONSUMO = c.CONSUMO_ADIF;
                        c.DIF_TOTAL = (c.cnpr_adif + c.cpre_adif);
                    }



                    d.Add(c.CUPSREE, c);
                }

                


                foreach (KeyValuePair<string, EndesaEntity.facturacion.Adif_Factura> p in d)
                {
                    if (diferencias)
                    {
                        if ((p.Value.CONSUMO_ADIF != p.Value.CONSUMO_SCE) || 
                            ((p.Value.cnpr_adif + p.Value.cpre_adif) != p.Value.cnpr_endesa))
                            lista.Add(p.Value);
                    }
                    else
                        lista.Add(p.Value);
                }



                Cursor.Current = Cursors.Default;

                if (lista.Count == 0)
                {
                    MessageBox.Show("La consulta que ha realizado no devuelve ningún resultado.",
                    "La consulta no devuelte datos.",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                }

                return lista;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la importación de CUPS",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

                return null;
            }

        }

        private Dictionary<string, EndesaEntity.facturacion.Factura> Facturas_BI(DateTime ffactdes, DateTime ffacthas, string cupsree)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.facturacion.Factura> d =
                new Dictionary<string, EndesaEntity.facturacion.Factura>();

            bool firstOnly = true;

            List<string> lista_conceptos_cnpr = p.GetValue("cnpr").Split(';').ToList();
            List<string> lista_conceptos_cpre = p.GetValue("cpre").Split(';').ToList();


            strSql = "SELECT f.CUPSREE, f.CREFEREN, f.SECFACTU, f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " SUM(f.VCUOVAFA) VCUOVAFA, f.IFACTURA, f.ISE, f.IVA,"
                    + " tf.descripcion as TFACTURA, f.TESTFACT"                    
                    + " FROM fact.fo f"
                    + " LEFT OUTER JOIN fact.fo_p_tipos_factura tf on"
                    + " tf.id_tipo_factura = f.TFACTURA"
                    + " WHERE f.CNIFDNIC = 'Q2802152E' AND"
                    + " (f.FFACTDES >= '" + ffactdes.ToString("yyyy-MM-dd") + "' AND"
                    + " f.FFACTHAS <= '" + ffacthas.ToString("yyyy-MM-dd") + "')";

            if (cupsree != null)
                strSql += " AND SUBSTR(f.CUPSREE, 1, 20) = '" + cupsree + "'";

            strSql += " group by f.CUPSREE, MONTH(f.ffactdes)";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.facturacion.Factura c = new EndesaEntity.facturacion.Factura();

                if (r["CUPSREE"] != System.DBNull.Value)
                    c.cupsree = Convert.ToString(r["CUPSREE"]).Substring(0, 20);

                if (r["CREFEREN"] != System.DBNull.Value)
                    c.creferen = Convert.ToInt64(r["CREFEREN"]);

                if (r["SECFACTU"] != System.DBNull.Value)
                    c.secfactu = Convert.ToInt32(r["SECFACTU"]);

                if (r["FFACTURA"] != System.DBNull.Value)
                    c.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                if (r["FFACTDES"] != System.DBNull.Value)
                    c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                if (r["FFACTHAS"] != System.DBNull.Value)
                    c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);

                if (r["VCUOVAFA"] != System.DBNull.Value)
                    c.vcuovafa = Convert.ToInt32(r["VCUOVAFA"]);

                if (r["TFACTURA"] != System.DBNull.Value)
                    c.tfactura_desc = Convert.ToString(r["TFACTURA"]);

                if (r["TESTFACT"] != System.DBNull.Value)
                    c.testfact = Convert.ToString(r["TESTFACT"]);

                if (r["IFACTURA"] != System.DBNull.Value)
                    c.ifactura = Convert.ToDouble(r["IFACTURA"]);               


                EndesaEntity.facturacion.Factura o;
                if (!d.TryGetValue(c.cupsree, out o))
                    d.Add(c.cupsree, c);
                
            }
            db.CloseConnection();

            strSql = "SELECT f.CUPSREE, f.CREFEREN, f.SECFACTU, f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " SUM(f.VCUOVAFA) VCUOVAFA, f.IFACTURA, f.ISE, f.IVA,"
                    + " tf.descripcion as TFACTURA, f.TESTFACT,"
                    + " t.TCONFAC, SUM(t.ICONFAC) ICONFAC"
                    + " FROM fact.fo f"
                    + " LEFT OUTER JOIN fact.fo_tcon t ON"
                    + " t.CREFEREN = f.CREFEREN AND"
                    + " t.SECFACTU = f.SECFACTU AND"
                    + " t.TESTFACT = f.TESTFACT AND"
                    + " t.TCONFAC IN (";
            
            foreach(string p in lista_conceptos_cnpr)
            {
                if (firstOnly)
                {
                    strSql += p;
                    firstOnly = false;
                }else
                    strSql += "," + p;

            }

            foreach (string p in lista_conceptos_cpre)
            {
                if (firstOnly)
                {
                    strSql += p;
                    firstOnly = false;
                }
                else
                    strSql += "," + p;

            }


            strSql += ") LEFT OUTER JOIN fact.fo_p_tipos_factura tf on"
                    + " tf.id_tipo_factura = f.TFACTURA"
                    + " WHERE f.CNIFDNIC = 'Q2802152E' AND"
                    + " (f.FFACTDES >= '" + ffactdes.ToString("yyyy-MM-dd") + "' AND"
                    + " f.FFACTHAS <= '" + ffacthas.ToString("yyyy-MM-dd") + "')";

            if (cupsree != null)
                strSql += " AND SUBSTR(f.CUPSREE, 1, 20) = '" + cupsree + "'";

            strSql += " group by f.CUPSREE, MONTH(f.ffactdes)";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.facturacion.Factura c = new EndesaEntity.facturacion.Factura();

                if (r["CUPSREE"] != System.DBNull.Value)
                    c.cupsree = Convert.ToString(r["CUPSREE"]).Substring(0,20);

                if (r["CREFEREN"] != System.DBNull.Value)
                    c.creferen = Convert.ToInt64(r["CREFEREN"]);

                if (r["SECFACTU"] != System.DBNull.Value)
                    c.secfactu = Convert.ToInt32(r["SECFACTU"]);

                if (r["FFACTURA"] != System.DBNull.Value)
                    c.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                if (r["FFACTDES"] != System.DBNull.Value)
                    c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                if (r["FFACTHAS"] != System.DBNull.Value)
                    c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);

                if (r["VCUOVAFA"] != System.DBNull.Value)
                    c.vcuovafa = Convert.ToInt32(r["VCUOVAFA"]);

                if (r["TFACTURA"] != System.DBNull.Value)
                    c.tfactura_desc = Convert.ToString(r["TFACTURA"]);

                if (r["TESTFACT"] != System.DBNull.Value)
                    c.testfact = Convert.ToString(r["TESTFACT"]);

                if (r["IFACTURA"] != System.DBNull.Value)
                    c.ifactura = Convert.ToDouble(r["IFACTURA"]);

                if (r["ICONFAC"] != System.DBNull.Value)
                    c.cnpr = Convert.ToDouble(r["ICONFAC"]);


                EndesaEntity.facturacion.Factura o;
                if (d.TryGetValue(c.cupsree, out o))
                    o.cnpr = c.cnpr;




            }
            db.CloseConnection();
            return d;
        }

        private Dictionary<string, EndesaEntity.facturacion.Adif_Factura> Facturas_ADIF(DateTime ffactdes, DateTime ffacthas)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.facturacion.Adif_Factura> d =
                new Dictionary<string, EndesaEntity.facturacion.Adif_Factura>();


            strSql = "SELECT m.cups20, m.lote, m.mes, sum(m.energia_consumida) AS energia_consumida,"
                    + " SUM(m.energia_facturada) energia_facturada, sum(m.cnpr) cnpr, sum(m.cpre) cpre"
                    + " FROM adif_ficheros_facturas_meses m"
                    + " WHERE m.mes >= " + ffactdes.ToString("yyyyMM") + " and"
                    + " m.mes <= " + ffacthas.ToString("yyyyMM")
                    + " group by m.cups20, m.mes";

           

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.facturacion.Adif_Factura c = new EndesaEntity.facturacion.Adif_Factura();

                if (r["cups20"] != System.DBNull.Value)
                    c.CUPSREE = r["cups20"].ToString();               

                if (r["lote"] != System.DBNull.Value)
                    c.LOTE = Convert.ToInt32(r["lote"]);

                if (r["energia_consumida"] != System.DBNull.Value)
                    c.CONSUMO_ADIF = Convert.ToInt32(r["energia_consumida"]);

                if (r["cnpr"] != System.DBNull.Value)
                    c.cnpr_adif = Convert.ToDouble(r["cnpr"]);

                if (r["cpre"] != System.DBNull.Value)
                    c.cpre_adif = Convert.ToDouble(r["cpre"]);

                EndesaEntity.facturacion.Adif_Factura o;
                if (!d.TryGetValue(c.CUPSREE, out o))                    
                    d.Add(c.CUPSREE, c);
               

            }
            db.CloseConnection();
            return d;
        }

        public  void ExportExcel()
        {
            int c = 1;
            int f = 1;
            SaveFileDialog save;
            try
            {

                if(lista_facturas.Count > 0)
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
                        var headerCells = workSheet.Cells[1, 1, 1, 20];
                        var headerFont = headerCells.Style.Font;

                        workSheet.View.FreezePanes(2, 1);

                        headerFont.Bold = true;
                        workSheet.Cells[f, c].Value = "CUPS20"; c++;
                        workSheet.Cells[f, c].Value = "LOTE"; c++;
                        workSheet.Cells[f, c].Value = "Med. baja"; c++;
                        workSheet.Cells[f, c].Value = "Dev. energía"; c++;
                        workSheet.Cells[f, c].Value = "Cierres Ener."; c++;
                        workSheet.Cells[f, c].Value = "Existe factura ADIF"; c++;
                        workSheet.Cells[f, c].Value = "Existe factura ENDESA"; c++;
                        workSheet.Cells[f, c].Value = "CREFEREN"; c++;
                        workSheet.Cells[f, c].Value = "SECFACTU"; c++;
                        workSheet.Cells[f, c].Value = "F. FACTURA"; c++;
                        workSheet.Cells[f, c].Value = "F. CONSUMO DESDE"; c++;
                        workSheet.Cells[f, c].Value = "F. CONSUMO HASTA"; c++;
                        workSheet.Cells[f, c].Value = "T. FACTURA"; c++;
                        workSheet.Cells[f, c].Value = "TESTFACT"; c++;
                        workSheet.Cells[f, c].Value = "CONSUMO ADIF"; c++;
                        workSheet.Cells[f, c].Value = "CONSUMO ENDESA"; c++;
                        workSheet.Cells[f, c].Value = "DIF CONSUMO"; c++;
                        workSheet.Cells[f, c].Value = "CNPR ADIF"; c++;
                        workSheet.Cells[f, c].Value = "CNPR ENDESA"; c++;
                        workSheet.Cells[f, c].Value = "DIF TOTAL"; c++;

                        foreach (EndesaEntity.facturacion.Adif_Factura p in lista_facturas)
                        {
                            f++;
                            c = 1;
                            workSheet.Cells[f, c].Value = p.CUPSREE; c++;
                            workSheet.Cells[f, c].Value = p.LOTE; c++;

                            if(p.medida_en_baja)
                                workSheet.Cells[f, c].Value = "Sí"; 
                            else
                                workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            if (p.devolucion_de_energia)
                                workSheet.Cells[f, c].Value = "Sí";
                            else
                                workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            if (p.cierres_energia)
                                workSheet.Cells[f, c].Value = "Sí";
                            else
                                workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            if (p.existe_factura_adif)
                                workSheet.Cells[f, c].Value = "Sí";
                            else
                                workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            if (p.existe_factura_sce)
                                workSheet.Cells[f, c].Value = "Sí";
                            else
                                workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;



                            if (p.FFACTURA > DateTime.MinValue)
                            {

                                workSheet.Cells[f, c].Value = p.CREFEREN; c++;
                                workSheet.Cells[f, c].Value = p.SECFACTU; c++;

                                workSheet.Cells[f, c].Value = p.FFACTURA;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                c++;
                                workSheet.Cells[f, c].Value = p.FFACTDES;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                c++;
                                workSheet.Cells[f, c].Value = p.FFACTHAS;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                c++;
                                workSheet.Cells[f, c].Value = p.de_tfactura; c++;
                                workSheet.Cells[f, c].Value = p.testfact; c++;

                            }
                            else
                            {
                                c++;
                                c++;
                                c++;
                                c++;
                                c++;
                                c++;
                                c++;
                            }

                            if (p.existe_factura_adif)
                            {
                                workSheet.Cells[f, c].Value = p.CONSUMO_ADIF;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            }
                           
                            c++;

                            if (p.FFACTURA > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.CONSUMO_SCE;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            }
                            c++;

                            workSheet.Cells[f, c].Value = p.DIF_CONSUMO;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            c++;

                            if (p.existe_factura_adif)
                            {
                                workSheet.Cells[f, c].Value = p.cnpr_adif;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            }
                            
                            c++;

                            if (p.FFACTURA > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.cnpr_endesa;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            }
                            c++;

                            workSheet.Cells[f, c].Value = p.DIF_TOTAL;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;
                        }


                        var allCells = workSheet.Cells[1, 1, f, 20];
                        workSheet.Cells["A1:T1"].AutoFilter = true;


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
                                   

                } // if (result == DialogResult.Yes)

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error en exportación Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
            }

        }

    }
}

