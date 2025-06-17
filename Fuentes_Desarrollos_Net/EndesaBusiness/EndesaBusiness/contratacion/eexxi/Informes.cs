using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.eexxi
{
    public class Informes
    {
        private contratacion.eexxi.Inventario inventario;        
        private EndesaBusiness.cartera.Cartera_SalesForce salesforce_cartera;
        private deuda.Deuda deuda;
        private utilidades.Param param;
        private Dictionary<string, EndesaEntity.contratacion.Inventario_Tabla> dic;


        public Informes(contratacion.eexxi.Inventario inv)
        {
            try
            {

                inventario = inv;
                Dictionary<string, string> dic_nifs = new Dictionary<string, string>();
                dic = new Dictionary<string, EndesaEntity.contratacion.Inventario_Tabla>();

                param = new utilidades.Param("eexxi_param", servidores.MySQLDB.Esquemas.CON);

                foreach (KeyValuePair<string, EndesaEntity.contratacion.Inventario_Tabla> p in inventario.dic_tmp_altas)
                {
                    string o;
                    if (!dic_nifs.TryGetValue(p.Value.nif, out o))
                        dic_nifs.Add(p.Value.nif, p.Value.nif);

                    EndesaEntity.contratacion.Inventario_Tabla oo;
                    if (!dic.TryGetValue(p.Key, out oo))
                        dic.Add(p.Key, p.Value);


                }

                foreach (KeyValuePair<string, EndesaEntity.contratacion.Inventario_Tabla> p in inventario.dic.Where(z => z.Value.vigente))
                {
                    string o;
                    if (!dic_nifs.TryGetValue(p.Value.nif, out o))
                        dic_nifs.Add(p.Value.nif, p.Value.nif);

                    EndesaEntity.contratacion.Inventario_Tabla oo;
                    if (!dic.TryGetValue(p.Key, out oo))
                        dic.Add(p.Key, p.Value);
                }

                
                salesforce_cartera = new cartera.Cartera_SalesForce(dic_nifs.Values.ToList());
                foreach (KeyValuePair<string, EndesaEntity.contratacion.Inventario_Tabla> p in dic.Where(z => z.Value.vigente))
                    //p.Value.carterizado = cartera.ExisteCartera(p.Value.nif);
                    p.Value.carterizado = salesforce_cartera.ExisteCartera(p.Value.nif);



                deuda = new deuda.Deuda(dic_nifs.Values.ToList());
                
                EnvioDatosVigentes();

                MessageBox.Show("Informe generado."
                    + System.Environment.NewLine
                    + System.Environment.NewLine
                    + "Revise el buzón: " + param.GetValue("buzon_mail", DateTime.Now, DateTime.Now)
                    + " para comprobar los correos generados.",
                        "Informes Mail",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                       "Error al generar los correos",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Error);
            }

        }
                
        private void EnvioDatosVigentes()
        {
            int f = 0;
            int c = 0;
            double max_pontencia = 0;
            string direccion = "";

            FileInfo fileInfo = new FileInfo(@"c:\Temp\ALTAS_EEXXI_" + DateTime.Now.ToString("yyyy_MM_dd_HHmmss") + ".xlsx");
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("INVENTARIO");

            var headerCells = workSheet.Cells[1, 1, 1, 18];
            var headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;
            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "CUPS"; c++;
            workSheet.Cells[f, c].Value = "CLIENTE"; c++;
            workSheet.Cells[f, c].Value = "CIF"; c++;
            workSheet.Cells[f, c].Value = "PROVINCIA"; c++;
            workSheet.Cells[f, c].Value = "MUNICIPIO"; c++;
            workSheet.Cells[f, c].Value = "FECHA ALTA"; c++;
            workSheet.Cells[f, c].Value = "FECHA BAJA"; c++;
            workSheet.Cells[f, c].Value = "territorial"; c++;            
            workSheet.Cells[f, c].Value = "Gestor"; c++;            
            workSheet.Cells[f, c].Value = "responsableZona"; c++;
            workSheet.Cells[f, c].Value = "responsableTerritorial"; c++;
            workSheet.Cells[f, c].Value = "MAX_POT_CTATADA"; c++;
            workSheet.Cells[f, c].Value = "NM_POT_CTATADA_1"; c++;
            workSheet.Cells[f, c].Value = "NM_POT_CTATADA_2"; c++;
            workSheet.Cells[f, c].Value = "NM_POT_CTATADA_3"; c++;
            workSheet.Cells[f, c].Value = "NM_POT_CTATADA_4"; c++;
            workSheet.Cells[f, c].Value = "NM_POT_CTATADA_5"; c++;
            workSheet.Cells[f, c].Value = "NM_POT_CTATADA_6"; 

            for (int i = 1; i <= c; i++)
            {
                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }


            #region Inventario Carterizado
            foreach (KeyValuePair<string, EndesaEntity.contratacion.Inventario_Tabla> p in dic.OrderBy(z => z.Value.nif).OrderBy(z => z.Key))
            {

                //direccion = cartera.Direccion(p.Value.nif);
                direccion = salesforce_cartera.Direccion(p.Value.nif);

                f++;
                c = 1;
                workSheet.Cells[f, c].Value = p.Value.cups22; c++;
                workSheet.Cells[f, c].Value = p.Value.razon_social; c++;
                workSheet.Cells[f, c].Value = p.Value.nif; c++;
                workSheet.Cells[f, c].Value = p.Value.cod_provincia; c++;
                workSheet.Cells[f, c].Value = p.Value.descripcion_poblacion; c++;

                workSheet.Cells[f, c].Value = p.Value.fecha_alta;
                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                c++;

                if (p.Value.fecha_baja > DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = p.Value.fecha_baja;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                //if (cartera.ExisteCartera(p.Value.nif))
                if (salesforce_cartera.ExisteCartera(p.Value.nif))
                {
                    workSheet.Cells[f, c].Value = salesforce_cartera.posicion; c++;
                    workSheet.Cells[f, c].Value = salesforce_cartera.gestor; c++;
                    workSheet.Cells[f, c].Value = salesforce_cartera.responsable_zona; c++;
                    workSheet.Cells[f, c].Value = salesforce_cartera.responsable_territorial; c++;
                }
                else
                    for (int i = 0; i < 4; i++)
                        c++;

                max_pontencia = 0;
                for (int i = 0; i < 6; i++)
                    if (p.Value.potencias[i] > max_pontencia)
                        max_pontencia = p.Value.potencias[i];

                workSheet.Cells[f, c].Value = max_pontencia;
                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                for (int i = 0; i < 6; i++)
                {
                    if (p.Value.potencias[i] > 0)
                    {
                        workSheet.Cells[f, c].Value = p.Value.potencias[i];
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    }
                    else
                    {
                        c++;
                    }
                }

                
                for (int i = 1; i <= c; i++)
                {

                    // Incidencias
                    if (p.Value.tiene_incidencia)
                    {
                        workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                           System.Drawing.Color.FromArgb(
                           Convert.ToInt32(param.GetValue("rgb_1_altas_incidencia", DateTime.Now, DateTime.Now)),
                           Convert.ToInt32(param.GetValue("rgb_2_altas_incidencia", DateTime.Now, DateTime.Now)),
                           Convert.ToInt32(param.GetValue("rgb_3_altas_incidencia", DateTime.Now, DateTime.Now))));
                    }

                    // BAJAS
                    if (p.Value.fecha_baja > DateTime.MinValue && p.Value.informado_alta == DateTime.MinValue)
                    {
                        workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                            System.Drawing.Color.FromArgb(
                            Convert.ToInt32(param.GetValue("rgb_1_bajas", DateTime.Now, DateTime.Now)),
                            Convert.ToInt32(param.GetValue("rgb_2_bajas", DateTime.Now, DateTime.Now)),
                            Convert.ToInt32(param.GetValue("rgb_3_bajas", DateTime.Now, DateTime.Now))));
                    }
                        
                    
                    else if(p.Value.informado_alta == DateTime.MinValue)
                    {
                        workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
                           System.Drawing.Color.FromArgb(
                           Convert.ToInt32(param.GetValue("rgb_1_altas", DateTime.Now, DateTime.Now)),
                           Convert.ToInt32(param.GetValue("rgb_2_altas", DateTime.Now, DateTime.Now)),
                           Convert.ToInt32(param.GetValue("rgb_3_altas", DateTime.Now, DateTime.Now))));
                    }
                       
                }

                
            }



            #endregion
            #region Inventario No Carterizado
            //foreach (KeyValuePair<string, EndesaEntity.contratacion.Inventario_Tabla> p in dic.OrderBy(z => z.Value.nif).OrderBy(z => z.Key))
            //{

            //    //direccion = cartera.Direccion(p.Value.nif);
            //    direccion = salesforce_cartera.Direccion(p.Value.nif);

            //    if (direccion == "")
            //    {
            //        f++;
            //        c = 1;
            //        workSheet.Cells[f, c].Value = p.Value.cups22; c++;
            //        workSheet.Cells[f, c].Value = p.Value.razon_social; c++;
            //        workSheet.Cells[f, c].Value = p.Value.nif; c++;
            //        workSheet.Cells[f, c].Value = p.Value.cod_provincia; c++;
            //        workSheet.Cells[f, c].Value = p.Value.descripcion_poblacion; c++;

            //        workSheet.Cells[f, c].Value = p.Value.fecha_alta;
            //        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
            //        c++;

            //        if (p.Value.fecha_baja > DateTime.MinValue)
            //        {
            //            workSheet.Cells[f, c].Value = p.Value.fecha_baja;
            //            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
            //        }
            //        c++;

            //            //if (cartera.ExisteCartera(p.Value.nif))
            //        if (salesforce_cartera.ExisteCartera(p.Value.nif))
            //        {

            //                //workSheet.Cells[f, c].Value = cartera.descResponsableTerritorial; c++;
            //                //workSheet.Cells[f, c].Value = cartera.nombreGestor; c++;
            //                //workSheet.Cells[f, c].Value = cartera.apellido1Gestor; c++;
            //                //workSheet.Cells[f, c].Value = cartera.apellido2Gestor; c++;
            //                //workSheet.Cells[f, c].Value = cartera.responsableZona; c++;
            //                //workSheet.Cells[f, c].Value = cartera.responsableTerritorial; c

            //                workSheet.Cells[f, c].Value = salesforce_cartera.descResponsableTerritorial; c++;
            //                workSheet.Cells[f, c].Value = salesforce_cartera.nombreGestor; c++;                            
            //                workSheet.Cells[f, c].Value = salesforce_cartera.responsableZona; c++;
            //                workSheet.Cells[f, c].Value = salesforce_cartera.responsableTerritorial; c++;


            //            }
            //            else
            //            for (int i = 0; i < 4; i++)
            //                c++;

            //        max_pontencia = 0;
            //        for (int i = 0; i < 6; i++)
            //            if (p.Value.potencias[i] > max_pontencia)
            //                max_pontencia = p.Value.potencias[i];

            //        workSheet.Cells[f, c].Value = max_pontencia;
            //        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
            //        for (int i = 0; i < 6; i++)
            //        {
            //            if (p.Value.potencias[i] > 0)
            //            {
            //                workSheet.Cells[f, c].Value = p.Value.potencias[i];
            //                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
            //            }
            //            else
            //            {
            //                c++;
            //            }
            //        }

            //        if (p.Value.temporal)
            //            for (int i = 1; i <= c; i++)
            //            {
            //                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            //                // BAJAS
            //                if (p.Value.fecha_baja > DateTime.MinValue)
            //                    workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
            //                        System.Drawing.Color.FromArgb(
            //                        Convert.ToInt32(param.GetValue("rgb_1_bajas", DateTime.Now, DateTime.Now)),
            //                        Convert.ToInt32(param.GetValue("rgb_2_bajas", DateTime.Now, DateTime.Now)),
            //                        Convert.ToInt32(param.GetValue("rgb_3_bajas", DateTime.Now, DateTime.Now))));
            //                else if (p.Value.tiene_incidencia)
            //                    workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
            //                        System.Drawing.Color.FromArgb(
            //                        Convert.ToInt32(param.GetValue("rgb_1_altas_incidencia", DateTime.Now, DateTime.Now)),
            //                        Convert.ToInt32(param.GetValue("rgb_2_altas_incidencia", DateTime.Now, DateTime.Now)),
            //                        Convert.ToInt32(param.GetValue("rgb_3_altas_incidencia", DateTime.Now, DateTime.Now))));
            //                else
            //                    workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(
            //                        System.Drawing.Color.FromArgb(
            //                        Convert.ToInt32(param.GetValue("rgb_1_altas", DateTime.Now, DateTime.Now)),
            //                        Convert.ToInt32(param.GetValue("rgb_2_altas", DateTime.Now, DateTime.Now)),
            //                        Convert.ToInt32(param.GetValue("rgb_3_altas", DateTime.Now, DateTime.Now))));
            //            }

            //    }
            //}

            var allCells = workSheet.Cells[1, 1, f, 18];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:R1"].AutoFilter = true;
            allCells.AutoFitColumns();


            #endregion

            #region Obligaciones

            

            workSheet = excelPackage.Workbook.Worksheets.Add("OBLIGACIONES");
            headerCells = workSheet.Cells[1, 1, 1, 14];
            headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;

            workSheet.Cells[f, c].Value = "CIF"; c++;
            workSheet.Cells[f, c].Value = "Cliente"; c++;
            workSheet.Cells[f, c].Value = "CUPS"; c++;
            workSheet.Cells[f, c].Value = "CFACTURA"; c++;
            workSheet.Cells[f, c].Value = "Importe"; c++;
            workSheet.Cells[f, c].Value = "fLimitePago"; c++;
            workSheet.Cells[f, c].Value = "fFact"; c++;
            workSheet.Cells[f, c].Value = "fDesde"; c++;
            workSheet.Cells[f, c].Value = "fHasta"; c++;
            workSheet.Cells[f, c].Value = "TFACTURA"; c++;
            workSheet.Cells[f, c].Value = "AAJJ"; c++;
            workSheet.Cells[f, c].Value = "AGRECOBRO"; c++;
            workSheet.Cells[f, c].Value = "PC"; c++;
            workSheet.Cells[f, c].Value = "FC"; c++;

            foreach (KeyValuePair<string, List<EndesaEntity.cobros.Deuda_Tabla>> p in deuda.dic)
            {

                    //direccion = cartera.Direccion(p.Key);
                direccion = salesforce_cartera.Direccion(p.Key);

                for (int i = 0; i < p.Value.Count; i++)
                {
                    c = 1;
                    f++;
                    workSheet.Cells[f, c].Value = p.Value[i].nif; c++;
                    workSheet.Cells[f, c].Value = p.Value[i].dapersoc; c++;
                    workSheet.Cells[f, c].Value = p.Value[i].cups22; c++;
                    workSheet.Cells[f, c].Value = p.Value[i].cfactura; c++;
                    workSheet.Cells[f, c].Value = p.Value[i].importe_obligacion;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    if (p.Value[i].fecha_limite_pago > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value[i].fecha_limite_pago;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.Value[i].fecha_factura > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value[i].fecha_factura;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.Value[i].periodo_factura_desde > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value[i].periodo_factura_desde;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.Value[i].periodo_factura_hasta > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value[i].periodo_factura_hasta;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    workSheet.Cells[f, c].Value = p.Value[i].tipo_factura; c++;
                    workSheet.Cells[f, c].Value = p.Value[i].aajj; c++;
                    workSheet.Cells[f, c].Value = p.Value[i].agrecobro; c++;
                    workSheet.Cells[f, c].Value = p.Value[i].pc; c++;
                    workSheet.Cells[f, c].Value = p.Value[i].fc; c++;

                }

            }

            #endregion

            allCells = workSheet.Cells[1, 1, f, c]; 
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:N1"].AutoFilter = true;
            allCells.AutoFitColumns();
            excelPackage.Save();

            //EnviaMail("ALTAS_EEXXI_EMPRESAS", fileInfo.FullName);            
            //EnviaMail("ALTAS_EEXXI_GGCC", fileInfo.FullName);
            EnviaMail("ALTAS_EEXXI", fileInfo.FullName);
        }


        private void EnviaMail(string plantilla, string rutaArchivo)
        {
           
            Enviar_Mails enviaMail = new Enviar_Mails();
            office.SendMail mail = new office.SendMail(param.GetValue("buzon_mail", DateTime.Now, DateTime.Now));


            Mail_Plantilla o;
            List<Mail_Destinatarios> oo;
            if (enviaMail.dic_plantillas.TryGetValue(plantilla, out o))
                if (enviaMail.dic_destinatarios.TryGetValue(plantilla, out oo))
                {

                    for (int i = 0; i < oo.Count; i++)
                        switch (oo[i].tipo_destinatario)
                        {
                            case "CC":
                                mail.cc.Add(oo[i].direccion_mail);
                                break;
                            case "Para":
                                mail.para.Add(oo[i].direccion_mail);
                                break;
                            case "CCO":
                                mail.cco.Add(oo[i].direccion_mail);
                                break;
                        }

                    // Añadimos a los RT desde la cartera de SalesForce
                    List<string> lista_rt = salesforce_cartera.ListaMail_Responsables_RT();
                    if(lista_rt.Count > 0)
                        for (int i = 0; i < lista_rt.Count; i++)
                            mail.para.Add(lista_rt[i]);

                    mail.sender = param.GetValue("buzon_mail", DateTime.Now, DateTime.Now);
                    mail.asunto = o.asunto;
                    mail.htmlCuerpo = o.cuerpo_mail;
                    mail.adjuntos.Add(rutaArchivo);

                    if (param.GetValue("enviar_mail", DateTime.Now, DateTime.Now) != "S")
                        mail.Save();
                    else
                        mail.Send();

                    // Update informado_alta
                    inventario.Update_Informado_Alta();
                }


        }
    }
}
