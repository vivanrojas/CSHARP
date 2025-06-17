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

namespace EndesaBusiness.eer
{
    public class InformeEERFacturas
    {
        public InformeEERFacturas()
        {

        }

        //private List<EndesaEntity.eer.InformeFacturacion> InformeFacturas(DateTime fd, DateTime fh)
        //{
        //    MySQLDB db;
        //    MySqlCommand command;
        //    MySqlDataReader r;
        //    string strSql;
        //    List<EndesaEntity.eer.InformeFacturacion> l = new List<EndesaEntity.eer.InformeFacturacion>();
        //    utilidades.Fechas utilFechas = new utilidades.Fechas();
        //    int dias = 0;
        //    int dias_anio = 0;


        //    dias_anio = utilFechas.EsAnioBisiesto(fd) ? 366 : 365;

        //    #region query
        //    if (fd >= new DateTime(2021, 06, 01))
        //    {
        //        strSql = "SELECT i.cups20 AS CUPS, i.nombre_cliente AS RAZON_SOCIAL, i.cif AS CIF,"
        //           + " NULL AS  DOMICILIO, NULL AS MUNICIPIO,"
        //           + " i.potencia_1, i.potencia_2, i.potencia_3, i.potencia_4, i.potencia_5, i.potencia_6,"
        //           + " i.tarifa_acceso, NULL AS REFERENCIA_CONTRATO,"
        //           + " f.codigo_factura, f.fecha_consumo_desde, f.fecha_consumo_hasta, f.fecha_factura AS FECHA_EMISION,"
        //           + " r.a_p1 AS CONSUMO_ACTIVA_PERIODO_1, r.a_p2 AS CONSUMO_ACTIVA_PERIODO_2,"
        //           + " r.a_p3 AS CONSUMO_ACTIVA_PERIODO_3, r.a_p4 AS CONSUMO_ACTIVA_PERIODO_4,"
        //           + " r.a_p5 AS CONSUMO_ACTIVA_PERIODO_5, r.a_p6 AS CONSUMO_ACTIVA_PERIODO_6,"
        //           + " r.r_p1 AS CONSUMO_REACTIVA_PERIODO_1, r.r_p2 AS CONSUMO_REACTIVA_PERIODO_2,"
        //           + " r.r_p3 AS CONSUMO_REACTIVA_PERIODO_3, r.r_p4 AS CONSUMO_REACTIVA_PERIODO_4,"
        //           + " r.r_p5 AS CONSUMO_REACTIVA_PERIODO_5, r.r_p6 AS CONSUMO_REACTIVA_PERIODO_6,"
        //           + " r.potmax_p1 AS POTENCIA_DEMANDADA_PERIODO_1,"
        //           + " r.potmax_p2 AS POTENCIA_DEMANDADA_PERIODO_2,"
        //           + " r.potmax_p3 AS POTENCIA_DEMANDADA_PERIODO_3,"
        //           + " r.potmax_p4 AS POTENCIA_DEMANDADA_PERIODO_4,"
        //           + " r.potmax_p5 AS POTENCIA_DEMANDADA_PERIODO_5,"
        //           + " r.potmax_p6 AS POTENCIA_DEMANDADA_PERIODO_6,"
        //           + " (SELECT SUM(TOTAL) FROM eer_facturasdetalle fd WHERE"
        //           + " fd.id_factura = f.id_factura AND"
        //           + " fd.producto = 'L85') AS COSTE_DE_POTENCIA,"
        //           + " (SELECT SUM(TOTAL) FROM eer_facturasdetalle fd WHERE"
        //           + " fd.id_factura = f.id_factura AND"
        //           + " fd.producto = 'REPA') AS EXCESOS_DE_POTENCIA,"
        //           + " f.termino_energia AS IMPORTE_ENERGIA,"
        //           + " (SELECT SUM(TOTAL) FROM eer_facturasdetalle fd WHERE"
        //           + " fd.id_factura = f.id_factura AND"
        //           + " fd.producto = 'L44') AS EXCESOS_DE_REACTIVA,"
        //           + " f.impuesto_electricidad AS IE,"
        //           + " f.impuesto_electricidad_reducido as ISE,"
        //           + " f.alquiler AS ALQUILER,"
        //           + " NULL AS FIANZA,"
        //           + " NULL AS DERECHOS_CONTRATACION,"
        //           + " f.base_imponible AS IMPORTE_SIN_IVA,"
        //           + " f.iva AS IVA,"
        //           + " f.total_factura"
        //           + " FROM eer_inventario i"
        //           + " LEFT OUTER JOIN eer_facturas f ON"
        //           + " f.cups20 = i.cups20"
        //           + " LEFT OUTER JOIN eer_resumen_medida r ON"
        //           + " r.cups20 = f.cups20 AND"
        //           + " r.fd = f.fecha_consumo_desde AND"
        //           + " r.fh = f.fecha_consumo_hasta"
        //           //+ " where (i.fecha_inicio <= '" + fd.ToString("yyyy-MM-dd") + "' AND"
        //           //+ " i.fecha_fin >= '" + fh.ToString("yyyy-MM-dd") + "') and"
        //           + " where "
        //           + " (f.fecha_consumo_desde >= '" + fd.ToString("yyyy-MM-dd") + "' and"
        //           + " f.fecha_consumo_hasta <= '" + fh.ToString("yyyy-MM-dd") + "')"
        //           + " and f.codigo_factura is not null";
        //    }
        //    else
        //    {
        //        strSql = "SELECT i.cups20 AS CUPS, i.nombre_cliente AS RAZON_SOCIAL, i.cif AS CIF,"
        //           + " NULL AS  DOMICILIO, NULL AS MUNICIPIO,"
        //           + " i.potencia_1, i.potencia_2, i.potencia_3, i.potencia_4, i.potencia_5, i.potencia_6,"
        //           + " i.tarifa_acceso, NULL AS REFERENCIA_CONTRATO,"
        //           + " f.codigo_factura, f.fecha_consumo_desde, f.fecha_consumo_hasta, f.fecha_factura AS FECHA_EMISION,"
        //           + " r.a_p1 AS CONSUMO_ACTIVA_PERIODO_1, r.a_p2 AS CONSUMO_ACTIVA_PERIODO_2,"
        //           + " r.a_p3 AS CONSUMO_ACTIVA_PERIODO_3, r.a_p4 AS CONSUMO_ACTIVA_PERIODO_4,"
        //           + " r.a_p5 AS CONSUMO_ACTIVA_PERIODO_5, r.a_p6 AS CONSUMO_ACTIVA_PERIODO_6,"
        //           + " r.r_p1 AS CONSUMO_REACTIVA_PERIODO_1, r.r_p2 AS CONSUMO_REACTIVA_PERIODO_2,"
        //           + " r.r_p3 AS CONSUMO_REACTIVA_PERIODO_3, r.r_p4 AS CONSUMO_REACTIVA_PERIODO_4,"
        //           + " r.r_p5 AS CONSUMO_REACTIVA_PERIODO_5, r.r_p6 AS CONSUMO_REACTIVA_PERIODO_6,"
        //           + " r.potmax_p1 AS POTENCIA_DEMANDADA_PERIODO_1,"
        //           + " r.potmax_p2 AS POTENCIA_DEMANDADA_PERIODO_2,"
        //           + " r.potmax_p3 AS POTENCIA_DEMANDADA_PERIODO_3,"
        //           + " r.potmax_p4 AS POTENCIA_DEMANDADA_PERIODO_4,"
        //           + " r.potmax_p5 AS POTENCIA_DEMANDADA_PERIODO_5,"
        //           + " r.potmax_p6 AS POTENCIA_DEMANDADA_PERIODO_6,"
        //           + " (SELECT SUM(TOTAL) FROM eer_facturasdetalle fd WHERE"
        //           + " fd.id_factura = f.id_factura AND"
        //           + " fd.producto = 'L34') AS COSTE_DE_POTENCIA,"
        //           + " (SELECT SUM(TOTAL) FROM eer_facturasdetalle fd WHERE"
        //           + " fd.id_factura = f.id_factura AND"
        //           + " fd.producto = 'REPA') AS EXCESOS_DE_POTENCIA,"
        //           + " f.termino_energia AS IMPORTE_ENERGIA,"
        //           + " (SELECT SUM(TOTAL) FROM eer_facturasdetalle fd WHERE"
        //           + " fd.id_factura = f.id_factura AND"
        //           + " fd.producto = 'L44') AS EXCESOS_DE_REACTIVA,"
        //           + " f.impuesto_electricidad AS IE,"
        //           + " f.impuesto_electricidad_reducido as ISE,"
        //           + " f.alquiler AS ALQUILER,"
        //           + " NULL AS FIANZA,"
        //           + " NULL AS DERECHOS_CONTRATACION,"
        //           + " f.base_imponible AS IMPORTE_SIN_IVA,"
        //           + " f.iva AS IVA,"
        //           + " f.total_factura"
        //           + " FROM eer_inventario i"
        //           + " LEFT OUTER JOIN eer_facturas f ON"
        //           + " f.cups20 = i.cups20"
        //           + " LEFT OUTER JOIN eer_resumen_medida r ON"
        //           + " r.cups20 = f.cups20 AND"
        //           + " r.fd = f.fecha_consumo_desde AND"
        //           + " r.fh = f.fecha_consumo_hasta"
        //           + " where (f.fecha_consumo_desde >= '" + fd.ToString("yyyy-MM-dd") + "' and"
        //           + " f.fecha_consumo_hasta <= '" + fh.ToString("yyyy-MM-dd") + "')"
        //           + " and f.codigo_factura is not null";
        //    }


        //    #endregion
        //    db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
        //    command = new MySqlCommand(strSql, db.con);
        //    r = command.ExecuteReader();

        //    while (r.Read())
        //    {

        //        EndesaEntity.eer.InformeFacturacion c = new EndesaEntity.eer.InformeFacturacion();
        //        c.cups = r["CUPS"].ToString(); 
        //        c.razon_social = r["RAZON_SOCIAL"].ToString(); 
        //        c.cif = r["CIF"].ToString(); 
        //        if (r["DOMICILIO"] != System.DBNull.Value)
        //            c.domicilio = r["DOMICILIO"].ToString();

        //        if (r["MUNICIPIO"] != System.DBNull.Value)
        //            c.municipio = r["MUNICIPIO"].ToString();

        //        for (int i = 1; i <= 6; i++)
        //        {
        //            if (r["potencia_" + i] != System.DBNull.Value)                    
        //                c.potencia[i] = Convert.ToDouble(r["potencia_" + i]);

        //        }
        //        c.tarifa_acceso = r["tarifa_acceso"].ToString(); 

        //        if (r["REFERENCIA_CONTRATO"] != System.DBNull.Value)
        //            c.referencia_contratato = r["REFERENCIA_CONTRATO"].ToString();

        //        if (r["codigo_factura"] != System.DBNull.Value)
        //            c.codigo_factura = r["codigo_factura"].ToString();

        //        if (r["fecha_consumo_desde"] != System.DBNull.Value)                
        //            c.fecha_consumo_desde = Convert.ToDateTime(r["fecha_consumo_desde"]);                

        //        if (r["fecha_consumo_hasta"] != System.DBNull.Value)                
        //            c.fecha_consumo_hasta = Convert.ToDateTime(r["fecha_consumo_hasta"]);                

        //        if (r["FECHA_EMISION"] != System.DBNull.Value)                
        //            c.fecha_emision = Convert.ToDateTime(r["FECHA_EMISION"]);                                    

        //        for (int i = 1; i <= 6; i++)
        //        {
        //            if (r["CONSUMO_ACTIVA_PERIODO_" + i] != System.DBNull.Value)                    
        //                c.consumo_activa[i] = Convert.ToDouble(r["CONSUMO_ACTIVA_PERIODO_" + i]);

        //        }

        //        for (int i = 1; i <= 6; i++)
        //        {
        //            if (r["CONSUMO_REACTIVA_PERIODO_" + i] != System.DBNull.Value)                    
        //                c.consumo_reactiva[i] = Convert.ToDouble(r["CONSUMO_REACTIVA_PERIODO_" + i]);                       

        //        }

        //        for (int i = 1; i <= 6; i++)
        //        {
        //            if (r["POTENCIA_DEMANDADA_PERIODO_" + i] != System.DBNull.Value)                    
        //                c.potencia_demandada[i] = Convert.ToDouble(r["POTENCIA_DEMANDADA_PERIODO_" + i]);

        //        }
        //        if (r["COSTE_DE_POTENCIA"] != System.DBNull.Value)
        //        {
        //            dias = Convert.ToInt32((c.fecha_consumo_hasta - c.fecha_consumo_desde).TotalDays + 1);
        //            c.coste_de_potencia = Convert.ToDouble(r["COSTE_DE_POTENCIA"]);
        //            c.coste_de_potencia = (c.coste_de_potencia / dias_anio) * dias;
        //        }

        //        if (r["EXCESOS_DE_POTENCIA"] != System.DBNull.Value)
        //        {
        //            c.excesos_de_potencia = Convert.ToDouble(r["EXCESOS_DE_POTENCIA"]);                    
        //        }

        //        if (r["IMPORTE_ENERGIA"] != System.DBNull.Value)
        //        {
        //            c.importe_energia = Convert.ToDouble(r["IMPORTE_ENERGIA"]);                    
        //        }

        //        if (r["EXCESOS_DE_REACTIVA"] != System.DBNull.Value)
        //        {
        //            c.excesos_de_reactiva = Convert.ToDouble(r["EXCESOS_DE_REACTIVA"]);                    
        //        }

        //        if (r["IE"] != System.DBNull.Value)
        //        {
        //            c.impuesto_electrico = Convert.ToDouble(r["IE"]);                
        //        }

        //        if (r["ISE"] != System.DBNull.Value)
        //        {
        //            c.impuesto_electrico_reducido = Convert.ToDouble(r["ISE"]);
        //        }

        //        if (r["ALQUILER"] != System.DBNull.Value)
        //        {
        //            c.alquiler_equipo = Convert.ToDouble(r["ALQUILER"]);                    
        //        }

        //        if (r["FIANZA"] != System.DBNull.Value)
        //        {
        //            c.fianza = Convert.ToDouble(r["FIANZA"]);                    
        //        }

        //        if (r["DERECHOS_CONTRATACION"] != System.DBNull.Value)
        //        {
        //            c.derechos_contratacion = Convert.ToDouble(r["DERECHOS_CONTRATACION"]);                    
        //        }

        //        if (r["IMPORTE_SIN_IVA"] != System.DBNull.Value)
        //        {
        //           c.importe_sin_iva = Convert.ToDouble(r["IMPORTE_SIN_IVA"]);                
        //        }

        //        if (r["IVA"] != System.DBNull.Value)
        //        {
        //            c.iva = Convert.ToDouble(r["IVA"]);                    
        //        }

        //        if (r["total_factura"] != System.DBNull.Value)
        //        {
        //            c.total_factura = Convert.ToDouble(r["total_factura"]);                    
        //        }

        //        l.Add(c);
        //    }
        //    db.CloseConnection();

        //    return l;
        //}


        private List<EndesaEntity.eer.InformeFacturacion> InformeFacturas(DateTime fd, DateTime fh)
        {

            medida.CurvaResumenEER cr = new medida.CurvaResumenEER(fd, fh);
            EndesaBusiness.medida.DatosPeajesEER peajesEER = new medida.DatosPeajesEER(fd, fh);
            EndesaEntity.DatosPeaje peaje;

            Inventario inventario = new Inventario(fd, fh);
            FacturasEER facturas = new FacturasEER(fd, fh);

            List<EndesaEntity.eer.InformeFacturacion> l = new List<EndesaEntity.eer.InformeFacturacion>();

            foreach (KeyValuePair<int, EndesaEntity.eer.Factura> p in facturas.dic)
            {
                

                EndesaEntity.eer.InformeFacturacion c = new EndesaEntity.eer.InformeFacturacion();
                c.cups = p.Value.cupsree;
                c.razon_social = p.Value.nombre_cliente;
                c.cif = p.Value.nif;
                c.domicilio = "";
                c.municipio = "";

                c.fecha_consumo_desde = p.Value.fecha_consumo_desde;
                c.fecha_consumo_hasta = p.Value.fecha_consumo_hasta;

                EndesaEntity.punto_suministro.PuntoSuministro ps =
                    inventario.GetPS(c.cif, c.cups, c.fecha_consumo_desde, c.fecha_consumo_hasta);
                

                for (int i = 1; i <= 6; i++)
                {
                    c.potencia[i] = ps.potecias_contratadas[i];
                }

                c.tarifa_acceso = ps.tarifa.tarifa;
                c.referencia_contratato = "";
                c.codigo_factura = p.Value.codigo_factura;
               
                c.fecha_emision = p.Value.fecha_factura;

                for (int i = 1; i <= 6; i++)
                {
                    c.consumo_activa[i] = p.Value.consumos_periodos_activa[i];
                    c.consumo_reactiva[i] = p.Value.consumos_periodos_reactiva[i];
                }

                EndesaEntity.medida.CurvaResumen ccrr = cr.GetCurvaResumen(c.cups, c.fecha_consumo_desde, c.fecha_consumo_hasta);
                if(ccrr != null)
                    for (int i = 1; i <= 6; i++)
                    {
                        c.consumo_activa[i] = ccrr.activa_periodo[i];
                        c.consumo_reactiva[i] = ccrr.reactiva_periodo[i];
                        c.potencia_demandada[i] = ccrr.potencias_maximas[i];
                    }
                else
                {
                    peaje = peajesEER.GetDatosPeaje(c.cups);
                    if(peaje != null)
                        for (int i = 1; i <= 6; i++)
                        {
                            c.consumo_activa[i] = peaje.activa[i];
                            c.consumo_reactiva[i] = peaje.reactiva[i];
                            c.potencia_demandada[i] = peaje.potmax[i];
                        }

                }                

                c.coste_de_potencia = p.Value.facturacion_potencia;
                c.excesos_de_potencia = p.Value.recargo_excesos;
                c.importe_energia = p.Value.termino_energia;
                c.excesos_de_reactiva = facturas.ImporteProducto(p.Value.id_factura, "L44");
                c.impuesto_electrico = p.Value.impuesto_electricidad;
                c.impuesto_electrico_reducido = p.Value.impuesto_electricidad_reducido;
                c.alquiler_equipo = p.Value.alquiler;
                c.fianza = 0;
                c.derechos_contratacion = 0;
                c.importe_sin_iva = p.Value.base_imponible;
                c.iva = p.Value.iva;
                c.total_factura = p.Value.total_factura;


                l.Add(c);
            }

            return l;
        }

        public void InformeExcel(DateTime fd, DateTime fh)
        {
            
            int c = 1;
            int f = 1;            
            SaveFileDialog save;
            try
            {
                List<EndesaEntity.eer.InformeFacturacion> lista = InformeFacturas(fd, fh);

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

                        foreach (EndesaEntity.eer.InformeFacturacion p in lista)
                        {


                            f++;
                            c = 1;
                            workSheet.Cells[f, c].Value = p.cups; c++;
                            workSheet.Cells[f, c].Value = p.razon_social; c++;
                            workSheet.Cells[f, c].Value = p.cif; c++;
                            workSheet.Cells[f, c].Value = p.domicilio; c++;
                            workSheet.Cells[f, c].Value = p.municipio; c++;

                            for (int i = 1; i <= 6; i++)
                            {
                                if (p.potencia[i] > 0)
                                {
                                    workSheet.Cells[f, c].Value = p.potencia[i];
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                }
                                c++;

                            }



                            workSheet.Cells[f, c].Value = p.tarifa_acceso; c++;
                            workSheet.Cells[f, c].Value = p.referencia_contratato; c++;

                            if (p.codigo_factura != null)
                                workSheet.Cells[f, c].Value = p.codigo_factura;
                            c++;

                            if (p.fecha_consumo_desde > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.fecha_consumo_desde;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;
                            if (p.fecha_consumo_hasta > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.fecha_consumo_hasta;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;
                            if (p.fecha_emision > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.fecha_emision;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;
                            for (int i = 1; i <= 6; i++)
                            {

                                workSheet.Cells[f, c].Value = p.consumo_activa[i];
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                c++;

                            }
                            for (int i = 1; i <= 6; i++)
                            {

                                workSheet.Cells[f, c].Value = p.consumo_reactiva[i];
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                c++;

                            }
                            for (int i = 1; i <= 6; i++)
                            {

                                workSheet.Cells[f, c].Value = p.potencia_demandada[i];
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                c++;

                            }


                            workSheet.Cells[f, c].Value = p.coste_de_potencia;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;


                            workSheet.Cells[f, c].Value = p.excesos_de_potencia;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;


                            workSheet.Cells[f, c].Value = p.importe_energia;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;


                            workSheet.Cells[f, c].Value = p.excesos_de_reactiva;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;

                            workSheet.Cells[f, c].Value = p.impuesto_electrico;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;

                            if(p.impuesto_electrico_reducido != 0) 
                            {
                                workSheet.Cells[f, c].Value = p.impuesto_electrico_reducido;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            }
                            
                            c++;


                            workSheet.Cells[f, c].Value = p.alquiler_equipo;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;


                            workSheet.Cells[f, c].Value = p.fianza;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;

                            workSheet.Cells[f, c].Value = p.derechos_contratacion;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";

                            c++;

                            workSheet.Cells[f, c].Value = p.importe_sin_iva;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;

                            workSheet.Cells[f, c].Value = p.iva;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;

                            workSheet.Cells[f, c].Value = p.total_factura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;





                        }

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
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "Exportación Excel",
               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
