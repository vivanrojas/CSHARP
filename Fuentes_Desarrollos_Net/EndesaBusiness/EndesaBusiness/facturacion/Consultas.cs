using EndesaBusiness.servidores;
using Microsoft.Graph;
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
    public class Consultas
    {
        public Consultas()
        {

        }


        public void FacturasSinCAP()
        {

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, Dictionary<int, EndesaEntity.consultas.FacturasSinCAP>> dic =
                new Dictionary<string, Dictionary<int, EndesaEntity.consultas.FacturasSinCAP>>();

            int tconfac = 0;
            double iconfac = 0;

            try
            {

                strSql = "SELECT f.CNIFDNIC, f.DAPERSOC, f.CCOUNIPS,"
                    + " f.CUPSREE, comp.CTARIFA,"
                    + " f.CREFEREN, f.SECFACTU, f.TESTFACT,"
                    + " f.CFACTURA,f.FFACTURA, f.FFACTDES, f.FFACTHAS, f.IFACTURA,"
                    + " tf.descripcion AS TFACTURA, t.TCONFAC, t.ICONFAC"
                    + " FROM cont.cont_comp_contratos_calendarios comp"
                    + " LEFT OUTER JOIN fact.fo f ON"
                    + " f.CCOUNIPS = comp.CCOUNIPS AND"
                    + " f.TESTFACT IN ('N','Y')"
                    + " INNER JOIN fact.fo_tcon t ON"
                    + " t.CREFEREN = f.CREFEREN AND"
                    + " t.SECFACTU = f.SECFACTU AND"
                    + " t.TESTFACT = f.TESTFACT"
                    + " LEFT OUTER JOIN fact.fo_p_tipos_factura tf ON"
                    + " tf.id_tipo_factura = f.TFACTURA"
                    + " WHERE comp.CCOMPOBJ = 'A01'"                    
                    + " AND(SUBSTR(comp.CCOUNIPS, 1, 1) <> 'U' AND SUBSTR(comp.CCOUNIPS, 1, 1) <> 'G' AND"
                    + " SUBSTR(comp.CCOUNIPS, 1, 3) <> 'XAA' AND SUBSTR(comp.CCOUNIPS, 1, 3) <> 'XCA')"
                    + " AND(f.FFACTDES >= '2022-06-01' AND f.FFACTHAS <= '2023-11-30')"                    
                    + " ORDER BY f.CNIFDNIC, f.CCOUNIPS, f.FFACTDES, f.SECFACTU";

                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.consultas.FacturasSinCAP c = new EndesaEntity.consultas.FacturasSinCAP();
                    c.cnifdnic = r["CNIFDNIC"].ToString();
                    c.dapersoc = r["DAPERSOC"].ToString();
                    c.ccounips = r["CCOUNIPS"].ToString();
                    c.cupsree = r["CUPSREE"].ToString();
                    c.ctarifa = r["CTARIFA"].ToString();
                    
                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                    if (r["FFACTDES"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                    if (r["FFACTHAS"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);

                    c.ifactura = Convert.ToDouble(r["IFACTURA"]);
                    c.tfactura = r["TFACTURA"].ToString();
                    c.secfactu = Convert.ToInt32(r["SECFACTU"]);
                    c.creferen = Convert.ToInt64(r["CREFEREN"]);
                    c.testfact = r["TESTFACT"].ToString();


                    tconfac = Convert.ToInt32(r["TCONFAC"]);
                    iconfac = Convert.ToDouble(r["ICONFAC"]);

                    Dictionary<int, EndesaEntity.consultas.FacturasSinCAP> d;
                    EndesaEntity.consultas.FacturasSinCAP o;
                    if (!dic.TryGetValue(c.ccounips + "_" + c.ffactdes.ToString("yyyyMMdd"), out d))
                    {

                        d = new Dictionary<int, EndesaEntity.consultas.FacturasSinCAP>();
                        c.dic.Add(tconfac, iconfac);
                        d.Add(c.secfactu, c);                                                
                        dic.Add(c.ccounips + "_" + c.ffactdes.ToString("yyyyMMdd"), d);
                    }
                    else
                    {
                        if (!d.TryGetValue(c.secfactu, out o))
                        {
                            c.dic.Add(tconfac, iconfac);
                            d.Add(c.secfactu, c);
                        }
                        else
                        {
                            if(!o.dic.TryGetValue(tconfac, out double v))
                                o.dic.Add(tconfac, iconfac);
                        }
                            
                    }
                    
                     
                }
                db.CloseConnection();
                FacturasSinCAP_Excel(dic);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        public void FacturasSinCAP_Excel(Dictionary<string, Dictionary<int, EndesaEntity.consultas.FacturasSinCAP>> dic)
        {
            int f = 1;
            int c = 1;

            string ruta_salida_archivo = @"c:\Temp\Borrar\FacturasPeninsula_sin_CAP.xlsx";

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            FileInfo fileInfo = new FileInfo(ruta_salida_archivo);
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);

            var workSheet = excelPackage.Workbook.Worksheets.Add("PS");
            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;

            workSheet.View.FreezePanes(2, 1);
            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "NIF"; c++;
            workSheet.Cells[f, c].Value = "CLIENTE"; c++;
            workSheet.Cells[f, c].Value = "CUPS13"; c++;
            workSheet.Cells[f, c].Value = "CUPSREE"; c++;            
            workSheet.Cells[f, c].Value = "TARIFA"; c++;
            workSheet.Cells[f, c].Value = "CREFEREN"; c++;
            workSheet.Cells[f, c].Value = "SECFACTU"; c++;
            workSheet.Cells[f, c].Value = "TESTFACT"; c++;
            workSheet.Cells[f, c].Value = "CFACTURA"; c++;
            workSheet.Cells[f, c].Value = "FFACTURA"; c++;
            workSheet.Cells[f, c].Value = "FFACTDES"; c++;
            workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
            workSheet.Cells[f, c].Value = "IFACTURA"; c++;
            workSheet.Cells[f, c].Value = "TFACTURA"; c++;

            foreach(KeyValuePair<string, Dictionary<int, EndesaEntity.consultas.FacturasSinCAP>> p in dic)
            {
                

                EndesaEntity.consultas.FacturasSinCAP pp = p.Value.Last().Value;

                double o;
                if(!pp.dic.TryGetValue(1258, out o))
                {
                    c = 1;
                    f++;

                    workSheet.Cells[f, c].Value = pp.cnifdnic; c++;
                    workSheet.Cells[f, c].Value = pp.dapersoc; c++;
                    workSheet.Cells[f, c].Value = pp.ccounips; c++;
                    workSheet.Cells[f, c].Value = pp.cupsree; c++;
                    workSheet.Cells[f, c].Value = pp.ctarifa; c++;
                    workSheet.Cells[f, c].Value = pp.creferen; c++;
                    workSheet.Cells[f, c].Value = pp.secfactu; c++;
                    workSheet.Cells[f, c].Value = pp.testfact; c++;

                    if (pp.cfactura != null)
                        workSheet.Cells[f, c].Value = pp.cfactura;
                    c++;

                    if(pp.ffactura > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = pp.ffactura;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (pp.ffactdes > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = pp.ffactdes;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (pp.ffacthas > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = pp.ffacthas;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    workSheet.Cells[f, c].Value = pp.ifactura;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    workSheet.Cells[f, c].Value = pp.tfactura; c++;

                    

                }

                
            }


            var allCells = workSheet.Cells[1, 1, f, c];
            headerCells = workSheet.Cells[1, 1, 1, c];
            headerFont = headerCells.Style.Font;
            headerFont.Bold = true;
            allCells = workSheet.Cells[1, 1, f, c];

            

            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:N1"].AutoFilter = true;
            allCells.AutoFitColumns();


            excelPackage.Save();


        }

    }
}
