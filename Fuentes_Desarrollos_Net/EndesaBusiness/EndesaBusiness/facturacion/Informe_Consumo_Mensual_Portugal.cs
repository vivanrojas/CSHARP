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
    public class Informe_Consumo_Mensual_Portugal
    {
        List<EndesaEntity.facturacion.Consumo_Medio_Mensual> lista { get; set; }
        public Informe_Consumo_Mensual_Portugal()
        {
            lista = new List<EndesaEntity.facturacion.Consumo_Medio_Mensual>();
        }

        public void Carga(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select Empresa, FECHA, TIPONEGOCIO, CONSUMO"
                    + " from fo_consumos_mes where "
                    + " fecha >= " + fd.ToString("yyyyMM") + " and"
                    + " fecha <= " + fh.ToString("yyyyMM");
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    EndesaEntity.facturacion.Consumo_Medio_Mensual c = new EndesaEntity.facturacion.Consumo_Medio_Mensual();
                    c.empresa = r["Empresa"].ToString();
                    c.aniomes = Convert.ToInt32(r["FECHA"]);
                    c.tipo_negocio = r["TIPONEGOCIO"].ToString();
                    c.consumo = Convert.ToInt64(r["CONSUMO"]);

                    lista.Add(c);
                }
                db.CloseConnection();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        public void InformeExcel(DateTime fd, DateTime fh)
        {

            int c = 1;
            int f = 1;
            SaveFileDialog save;
            try
            {
                Carga(fd, fh);

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
                        workSheet.Cells[f, c].Value = "CUPS"; c++;
                        workSheet.Cells[f, c].Value = "RAZÓN SOCIAL"; c++;
                        workSheet.Cells[f, c].Value = "CIF"; c++;
                        workSheet.Cells[f, c].Value = "DOMICILIO"; c++;
                        workSheet.Cells[f, c].Value = "MUNICIPIO"; c++;
                        for (int i = 1; i <= 6; i++)
                        {
                            workSheet.Cells[f, c].Value = "POTENCIA P" + i; c++;
                        }
                        workSheet.Cells[f, c].Value = "TARIFA ACCESO"; c++;
                        workSheet.Cells[f, c].Value = "REFERENCIA CONTRATO"; c++;
                        workSheet.Cells[f, c].Value = "CÓDIGO FACTURA"; c++;
                        workSheet.Cells[f, c].Value = "FECHA CONSUMO DESDE"; c++;
                        workSheet.Cells[f, c].Value = "FECHA CONSUMO HASTA"; c++;
                        workSheet.Cells[f, c].Value = "FECHA EMISIÓN"; c++;

                        for (int i = 1; i <= 6; i++)
                        {
                            workSheet.Cells[f, c].Value = "CONSUMO ACTIVA P" + i; c++;
                        }
                        for (int i = 1; i <= 6; i++)
                        {
                            workSheet.Cells[f, c].Value = "CONSUMO REACTIVA P" + i; c++;
                        }
                        for (int i = 1; i <= 6; i++)
                        {
                            workSheet.Cells[f, c].Value = "POTENCIA DEMANDADA P" + i; c++;
                        }

                        workSheet.Cells[f, c].Value = "COSTE DE POTENCIA"; c++;
                        workSheet.Cells[f, c].Value = "EXCESOS DE POTENCIA"; c++;
                        workSheet.Cells[f, c].Value = "IMPORTE ENERGÍA"; c++;
                        workSheet.Cells[f, c].Value = "EXCESOS DE REACTIVA"; c++;
                        workSheet.Cells[f, c].Value = "IMPUESTO ELÉCTRICO"; c++;
                        workSheet.Cells[f, c].Value = "IMPUESTO ELÉCTRICO REDUCIDO"; c++;
                        workSheet.Cells[f, c].Value = "ALQUILER EQUIPO"; c++;
                        workSheet.Cells[f, c].Value = "FIANZA"; c++;
                        workSheet.Cells[f, c].Value = "DERECHOS CONTRATACIÓN"; c++;
                        workSheet.Cells[f, c].Value = "IMPORTE SIN IVA"; c++;
                        workSheet.Cells[f, c].Value = "IVA"; c++;
                        workSheet.Cells[f, c].Value = "TOTAL FACTURA"; c++;

                        //foreach (EndesaEntity.facturacion.Consumo_Medio_Mensual p in lista)
                        //{


                        //    f++;
                        //    c = 1;
                        //    workSheet.Cells[f, c].Value = p.cups; c++;
                        //    workSheet.Cells[f, c].Value = p.razon_social; c++;
                        //    workSheet.Cells[f, c].Value = p.cif; c++;
                        //    workSheet.Cells[f, c].Value = p.domicilio; c++;
                        //    workSheet.Cells[f, c].Value = p.municipio; c++;

                        //    for (int i = 1; i <= 6; i++)
                        //    {
                        //        if (p.potencia[i] > 0)
                        //        {
                        //            workSheet.Cells[f, c].Value = p.potencia[i];
                        //            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        //        }
                        //        c++;

                        //    }



                        //    workSheet.Cells[f, c].Value = p.tarifa_acceso; c++;
                        //    workSheet.Cells[f, c].Value = p.referencia_contratato; c++;

                        //    if (p.codigo_factura != null)
                        //        workSheet.Cells[f, c].Value = p.codigo_factura;
                        //    c++;

                        //    if (p.fecha_consumo_desde > DateTime.MinValue)
                        //    {
                        //        workSheet.Cells[f, c].Value = p.fecha_consumo_desde;
                        //        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        //    }
                        //    c++;
                        //    if (p.fecha_consumo_hasta > DateTime.MinValue)
                        //    {
                        //        workSheet.Cells[f, c].Value = p.fecha_consumo_hasta;
                        //        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        //    }
                        //    c++;
                        //    if (p.fecha_emision > DateTime.MinValue)
                        //    {
                        //        workSheet.Cells[f, c].Value = p.fecha_emision;
                        //        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        //    }
                        //    c++;
                        //    for (int i = 1; i <= 6; i++)
                        //    {

                        //        workSheet.Cells[f, c].Value = p.consumo_activa[i];
                        //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        //        c++;

                        //    }
                        //    for (int i = 1; i <= 6; i++)
                        //    {

                        //        workSheet.Cells[f, c].Value = p.consumo_reactiva[i];
                        //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        //        c++;

                        //    }
                        //    for (int i = 1; i <= 6; i++)
                        //    {

                        //        workSheet.Cells[f, c].Value = p.potencia_demandada[i];
                        //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        //        c++;

                        //    }


                        //    workSheet.Cells[f, c].Value = p.coste_de_potencia;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //    c++;


                        //    workSheet.Cells[f, c].Value = p.excesos_de_potencia;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //    c++;


                        //    workSheet.Cells[f, c].Value = p.importe_energia;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //    c++;


                        //    workSheet.Cells[f, c].Value = p.excesos_de_reactiva;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //    c++;

                        //    workSheet.Cells[f, c].Value = p.impuesto_electrico;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //    c++;

                        //    if (p.impuesto_electrico_reducido != 0)
                        //    {
                        //        workSheet.Cells[f, c].Value = p.impuesto_electrico_reducido;
                        //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //    }

                        //    c++;


                        //    workSheet.Cells[f, c].Value = p.alquiler_equipo;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //    c++;


                        //    workSheet.Cells[f, c].Value = p.fianza;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //    c++;

                        //    workSheet.Cells[f, c].Value = p.derechos_contratacion;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";

                        //    c++;

                        //    workSheet.Cells[f, c].Value = p.importe_sin_iva;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //    c++;

                        //    workSheet.Cells[f, c].Value = p.iva;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //    c++;

                        //    workSheet.Cells[f, c].Value = p.total_factura;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //    c++;





                        //}

                        var allCells = workSheet.Cells[1, 1, f, 47];
                        workSheet.Cells["A1:AU1"].AutoFilter = true;


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
                else
                {
                    MessageBox.Show("No existen facturas registradas para el periodo de consumo:"
                        + System.Environment.NewLine
                        + fd.ToString("dd/MM/yyyy") + " al " + fh.ToString("dd/MM/yyyy"),
                        "Informe",
                         MessageBoxButtons.OK, MessageBoxIcon.Information);
                }


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Exportación Excel",
               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
