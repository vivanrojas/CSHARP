using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class CuadroMandoGASMedida
    {
        Dictionary<int, EndesaEntity.medida.CM_GAS_Cisterna> d_cisternas = new Dictionary<int, EndesaEntity.medida.CM_GAS_Cisterna>();
        Dictionary<int, EndesaEntity.medida.CM_Gas_Tubo> d_tubo = new Dictionary<int, EndesaEntity.medida.CM_Gas_Tubo>();
        public CuadroMandoGASMedida()
        {
            utilidades.Fechas fecha = new utilidades.Fechas();
            sigame.SIGAME sig = new sigame.SIGAME();
            sigame.GAS2000 clientes_gas2000 = new sigame.GAS2000();
            sigame.Albaranes_Clientes albaranes = new sigame.Albaranes_Clientes();
            sigame.DistribuidorasGasMySQL distribuidoras = new sigame.DistribuidorasGasMySQL();

            facturacion.TamGas tam = new facturacion.TamGas();
            //medida.FacturacionGas facturacionGas = new FacturacionGas();
            facturacion.FacturasGas facturacionGas = new facturacion.FacturasGas();


            DateTime ultimo_dia_habil = new DateTime();
            ultimo_dia_habil = fecha.UltimoDiaHabil();
            int ultimo_periodo_facturado = 0;
            int ultimoPeriodo = Convert.ToInt32(DateTime.Now.AddMonths(-1).ToString("yyyyMM"));


            facturacionGas.LanzaCargaUltimaFactura();

            foreach(KeyValuePair<string, EndesaEntity.sigame.ContratoGas> p in sig.dic)
            {
                clientes_gas2000.GetId_pto_suministro(p.Value.id_ps);

                ultimo_periodo_facturado = facturacionGas.UltimaFactura(p.Value.id_ps);

                if (!p.Value.es_cisterna)
                {

                    EndesaEntity.medida.CM_Gas_Tubo t = new EndesaEntity.medida.CM_Gas_Tubo();
                    t.fecha_alta = p.Value.fecha_inicio;
                    t.cups = p.Value.cupsree;
                    t.pais = p.Value.pais;
                    if (t.cups.Substring(0, 2) == "PT")
                        t.distribuidora = distribuidoras.GetDistribuidoraPT(t.cups);
                    else if (p.Value.distribuidora != null)
                        t.distribuidora = distribuidoras.GetDistribuidoraES(p.Value.distribuidora);
                    else
                        t.distribuidora = "";

                    t.nombre_cliente = p.Value.nombre_cliente;                    

                    if (clientes_gas2000.existe_id_pto_suministro)
                    {
                        
                        t.telemedida = clientes_gas2000.telemedida;
                        t.top = clientes_gas2000.top;
                        t.mes = Convert.ToInt32(clientes_gas2000.mes_medida);
                        t.area_pdte = Area_Pendiente(clientes_gas2000.mes_medida, ultimo_periodo_facturado.ToString());
                        t.motivo_pdte = clientes_gas2000.motivo_pendiente;
                    }

                    t.primer_mes_pdte = ultimo_periodo_facturado;
                    t.tam = tam.GetTam_ID_PS(p.Value.id_ps);

                    EndesaEntity.medida.CM_Gas_Tubo o;
                    if(!d_tubo.TryGetValue(p.Value.id_ps, out o))
                        d_tubo.Add(p.Value.id_ps,t);

                }
                else
                {
                    EndesaEntity.medida.CM_GAS_Cisterna c = new EndesaEntity.medida.CM_GAS_Cisterna();
                    c.nombre_cliente = clientes_gas2000.nombre_cliente;

                    c.fecha_alta = p.Value.fecha_inicio;
                    c.nombre_cliente = p.Value.nombre_cliente;
                    c.primer_mes_dpte = ultimo_periodo_facturado;
                    c.mes = Convert.ToInt32(clientes_gas2000.mes_medida);
                    c.area_pdte = Area_Pendiente(clientes_gas2000.mes_medida, ultimo_periodo_facturado.ToString());
                    c.tam = tam.GetTam_ID_PS(p.Value.id_ps);
                    c.codigo_slm = albaranes.Get_Codigo_SLM(p.Value.id_ps);

                    EndesaEntity.medida.CM_GAS_Cisterna o;
                    if(!d_cisternas.TryGetValue(p.Value.id_ps, out o))                    
                        d_cisternas.Add(p.Value.id_ps, c);
                    

                }

            }

            InformeExcel(@"c:\Temp\CMM_GAS" + DateTime.Now.ToString("yyyyMMdd_HHmmss")  +  ".xlsx");

        }
        public void InformeExcel(string fichero)
        {
            int f = 1;
            int c = 1;

            FileInfo fileInfo = new FileInfo(fichero);

            if (fileInfo.Exists)
                fileInfo.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("TUBO");

            var headerCells = workSheet.Cells[1, 1, 1, 28];
            var headerFont = headerCells.Style.Font;
            f = 1;

            headerFont.Bold = true;
            #region Cabecera_Excel

            workSheet.Cells[f, c].Value = "ID_PS";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "CUPS";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Fecha Alta";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "País";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Nombre";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Distribuidora";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "TM";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "TOP";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Primer mes pdte";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Último mes validado";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;


            workSheet.Cells[f, c].Value = "Área pdte";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Motivo pdte";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "TAM";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            #endregion

            foreach (KeyValuePair<int, EndesaEntity.medida.CM_Gas_Tubo> p in d_tubo)
            {
                c = 1;
                f++;

                workSheet.Cells[f, c].Value = p.Key; c++; 
                workSheet.Cells[f, c].Value = p.Value.cups; c++;
                workSheet.Cells[f, c].Value = p.Value.fecha_alta;
                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.Value.pais; c++;
                workSheet.Cells[f, c].Value = p.Value.nombre_cliente; c++;
                workSheet.Cells[f, c].Value = p.Value.distribuidora; c++;
                workSheet.Cells[f, c].Value = p.Value.telemedida ? "S" : "N"; c++;
                workSheet.Cells[f, c].Value = p.Value.top ? "S" : "N"; c++;
                workSheet.Cells[f, c].Value = p.Value.primer_mes_pdte; c++;
                workSheet.Cells[f, c].Value = p.Value.mes;c++;
                workSheet.Cells[f, c].Value = p.Value.area_pdte; c++;
                workSheet.Cells[f, c].Value = p.Value.motivo_pdte; c++;
                workSheet.Cells[f, c].Value = p.Value.tam; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
            }

            var allCells = workSheet.Cells[1, 1, f, 28];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:K1"].AutoFilter = true;
            allCells.AutoFitColumns();


            workSheet = excelPackage.Workbook.Worksheets.Add("CISTERNAS");
            headerCells = workSheet.Cells[1, 1, 1, 28];
            headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;

            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "ID_PS";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Fecha Alta";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Nombre";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Primer mes pdte";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Último mes validado";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;


            workSheet.Cells[f, c].Value = "Área pdte";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "TAM";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

            workSheet.Cells[f, c].Value = "Código SLM";
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;


            foreach (KeyValuePair<int, EndesaEntity.medida.CM_GAS_Cisterna> p in d_cisternas)
            {
                c = 1;
                f++;

                workSheet.Cells[f, c].Value = p.Key; c++;
                workSheet.Cells[f, c].Value = p.Value.fecha_alta;
                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.Value.nombre_cliente; c++;
                workSheet.Cells[f, c].Value = p.Value.primer_mes_dpte; c++;
                workSheet.Cells[f, c].Value = p.Value.mes; c++;
                workSheet.Cells[f, c].Value = p.Value.area_pdte; c++;
                workSheet.Cells[f, c].Value = p.Value.tam; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = p.Value.codigo_slm; c++;
            }

            allCells = workSheet.Cells[1, 1, f, 5];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:H1"].AutoFilter = true;
            allCells.AutoFitColumns();


            excelPackage.Save();


        }

        private string Area_Pendiente(string mes_medida, string primer_mes_pdte)
        {
            if (primer_mes_pdte == DateTime.Now.ToString("yyyyMM"))
                return "FACTURADO";
            else if (mes_medida == DateTime.Now.AddMonths(-1).ToString("yyyyMM"))
                return "Pdte. Facturación";
            else
                return "Pdte. Medida";
        }
    }
}
