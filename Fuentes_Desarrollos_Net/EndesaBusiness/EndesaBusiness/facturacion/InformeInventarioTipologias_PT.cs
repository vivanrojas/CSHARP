using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using MimeKit;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Renci.SshNet.Common;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.utilidades;
using EndesaEntity.contratacion;

namespace EndesaBusiness.facturacion
{
    public class InformeInventarioTipologias_PT
    {
        logs.Log ficheroLog;
        utilidades.Param param;


        EndesaBusiness.medida.TAM tam;



        // Agrupaciones
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_empresa;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_tarifa;        
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_territorio;

        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_agrupada;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_age;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_agora;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_passthough;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_passpool_periodo;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_passpool_subasta;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_tipo_contrato;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_gestion_atr;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_estado_contrato;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_revendedores;

        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_tipo_punto_producto;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_producto;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_tipo_tarifa;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_ciclo_p01_mt;
        Dictionary<string, EndesaEntity.Agrupacion> dic_agrup_exencion;

        Dictionary<DateTime, List<EndesaEntity.facturacion.IRF_Totales_Hist>> dic_totales;

        // Para el control de la ejecucion
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;

        // Exenciones
        Dictionary<string, EndesaEntity.contratacion.Exenciones> dic_exenciones;


        public InformeInventarioTipologias_PT()
        {

            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_InformeInventarioTipologias_PT");
            param = new utilidades.Param("informe_inventario_tipologias_pt_param", servidores.MySQLDB.Esquemas.CON);


            dic_agrup_empresa = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_tarifa = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_territorio = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_agrupada = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_age = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_agora = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_passthough = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_passpool_periodo = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_passpool_subasta = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_tipo_contrato = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_gestion_atr = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_estado_contrato = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_revendedores = new Dictionary<string, EndesaEntity.Agrupacion>();

            dic_agrup_tipo_punto_producto = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_producto = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_tipo_tarifa = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_ciclo_p01_mt = new Dictionary<string, EndesaEntity.Agrupacion>();
            dic_agrup_exencion = new Dictionary<string, EndesaEntity.Agrupacion>();


            ss_pp = new utilidades.Seguimiento_Procesos();

            dic_exenciones = new Dictionary<string, EndesaEntity.contratacion.Exenciones>();

        }

        public void Inventario_por_Tipologias()
        {
            EndesaBusiness.contratacion.ContratosPS_Portugal_BTN btn = new contratacion.ContratosPS_Portugal_BTN();
            EndesaBusiness.contratacion.ContratosPS_Portugal_BTE bte = new contratacion.ContratosPS_Portugal_BTE();
            EndesaBusiness.contratacion.ContratosPS_Portugal_MT mt = new contratacion.ContratosPS_Portugal_MT();
            


            bool firstOnly = true;
            int f = 1;
            int c = 1;

            EndesaEntity.Agrupacion o;


            // Totales para agrupaciones

            int total_agrupada = 0;
            int total_age = 0;
            int total_agora = 0;
            int total_passthough = 0;
            int total_passpool_horario = 0;
            int total_passpool_periodo = 0;
            int total_passpool_subasta = 0;
            int total_gestion_atr = 0;
            int total_revendedores = 0;
            int total_medido_baja = 0;
            int total_multipuntos = 0;
            double valor_tam = 0;
            int total_reg = 0;

            try
            {

                btn.CargaInventarioBTN();
                bte.CargaInventarioBTE();
                mt.CargaInventarioMT();

                // Carga MultiPuntos COMPOR
                EndesaBusiness.compor.Inventario_Multipuntos inventario_multipuntos
                    = new compor.Inventario_Multipuntos();


                tam = new medida.TAM();
                tam.CargaTAM();

                string ruta_salida_archivo = param.GetValue("Ubicacion_Informes")
                 + param.GetValue("Excel_prefijo_Tipologias") + "_"
                 + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";


                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                FileInfo fileInfo = new FileInfo(ruta_salida_archivo);
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);

                var workSheet = excelPackage.Workbook.Worksheets.Add("PS");
                var headerCells = workSheet.Cells[1, 1, 1, 25];
                var headerFont = headerCells.Style.Font;

                workSheet.View.FreezePanes(2, 1);

                headerFont.Bold = true;

                utilidades.Fechas utilfecha = new Fechas();


                #region Cabecera
                workSheet.Cells[f, c].Value = "CUPS"; c++;
                workSheet.Cells[f, c].Value = "CUPS22"; c++;
                workSheet.Cells[f, c].Value = "ESTADO CONTRATO";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;
                workSheet.Cells[f, c].Value = "F.ALTA"; c++;
                workSheet.Cells[f, c].Value = "F.BAJA"; c++;                
                workSheet.Cells[f, c].Value = "CIF"; c++;
                workSheet.Cells[f, c].Value = "RAZÓN SOCIAL"; c++;                
                workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                workSheet.Cells[f, c].Value = "TARIFA"; c++;                
                workSheet.Cells[f, c].Value = "PROVINCIA"; c++;
                workSheet.Cells[f, c].Value = "TERRITORIO"; c++;
                workSheet.Cells[f, c].Value = "CONTRATO EXT"; c++;
                workSheet.Cells[f, c].Value = "AGRUPADA"; c++;  
                
                workSheet.Cells[f, c].Value = "AGORA";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                
                //workSheet.Cells[f, c].Value = "GESTIÓN PROPIA ATR"; c++;
                workSheet.Cells[f, c].Value = "TAM";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;
                workSheet.Cells[f, c].Value = "ESTADO CONTRATO"; c++;
                workSheet.Cells[f, c].Value = "VERSIÓN"; c++;
                workSheet.Cells[f, c].Value = "F. PUESTA SERVICIO"; c++;
                workSheet.Cells[f, c].Value = "REVENDEDORES"; c++;
                // workSheet.Cells[f, c].Value = "ÚLTIMO CONVERTIDO"; c++;
                workSheet.Cells[f, c].Value = "MEDIDO EN BAJA"; c++;
                workSheet.Cells[f, c].Value = "MULTIPUNTO"; c++;
                #endregion

                foreach (KeyValuePair<string, EndesaEntity.facturacion.InventarioTipologias> p in btn.inventario)
                {
                                       
                    c = 1;
                    f++;

                    workSheet.Cells[f, c].Value = p.Value.cups13; c++;
                    workSheet.Cells[f, c].Value = p.Value.cups22; c++;

                    workSheet.Cells[f, c].Value = p.Value.estado_contrato;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                    if (p.Value.fecha_alta > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value.fecha_alta;  // F. alta
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.Value.fecha_baja > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value.fecha_baja;  // F. Baja
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;




                    workSheet.Cells[f, c].Value = p.Value.cif; c++;
                    workSheet.Cells[f, c].Value = p.Value.razon_social; c++;

                   

                    workSheet.Cells[f, c].Value = p.Value.empresa; c++;

                    if (!dic_agrup_empresa.TryGetValue(p.Value.empresa, out o))
                    {
                        o = new EndesaEntity.Agrupacion();
                        o.tipo = p.Value.empresa;
                        o.total = 1;
                        dic_agrup_empresa.Add(o.tipo, o);
                    }
                    else
                        o.total++;
                    
                    workSheet.Cells[f, c].Value = p.Value.tarifa; c++;

                    if (!dic_agrup_tarifa.TryGetValue(p.Value.tarifa, out o))
                    {
                        o = new EndesaEntity.Agrupacion();
                        o.tipo = p.Value.tarifa;
                        o.total = 1;
                        dic_agrup_tarifa.Add(o.tipo, o);
                    }
                    else
                        o.total++;


                    // workSheet.Cells[f, c].Value = p.Value.tipo_punto_suministro; c++;
                    workSheet.Cells[f, c].Value = p.Value.provincia; c++;
                    workSheet.Cells[f, c].Value = p.Value.territorio; c++;
                    workSheet.Cells[f, c].Value = p.Value.contrato_ext; c++;

                    // AGRUPADA
                    
                    workSheet.Cells[f, c].Value = "No";
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    
                    c++;    

                    // AGORA
                    workSheet.Cells[f, c].Value = ""; 
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                    // PASSTHOUGH	
                    workSheet.Cells[f, c].Value = ""; 
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                    // PASSPOOL HORARIO	
                    workSheet.Cells[f, c].Value = ""; 
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                    // PASSPOOL PERIODO	
                    workSheet.Cells[f, c].Value = "";  
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                    // PASSPOOL SUBASTA	
                    workSheet.Cells[f, c].Value = ""; 
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                    // TIPO CONTRATO	
                    workSheet.Cells[f, c].Value = ""; 
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                    // GESTIÓN PROPIA ATR	
                    //workSheet.Cells[f, c].Value = "No"; 
                    //workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;


                    if (p.Value.cups22 != null)
                    {
                        valor_tam = tam.GetTAM(p.Value.cups22);

                        // TAM
                        if (valor_tam != -1111111)
                        {
                            workSheet.Cells[f, c].Value = valor_tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        }
                        else
                        {
                            workSheet.Cells[f, c].Value = "Sin Facturas";
                        }
                        c++;
                    }
                    else
                        c++;

                    // ESTADO CONTRATO
                    workSheet.Cells[f, c].Value = p.Value.estado_contrato;  
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;


                    workSheet.Cells[f, c].Value = p.Value.version; c++; // VERSION

                    if (p.Value.fecha_puesta_servicio > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value.fecha_puesta_servicio;  // F. PUESTA SERVICIO
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;


                    workSheet.Cells[f, c].Value = "No";  // REVENDEDORES
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    //workSheet.Cells[f, c].Value = "No";  // ULTIMO CONVERTIDO
                    //workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //c++;

                    workSheet.Cells[f, c].Value = "No";  // MEDIDO EN BAJA
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;


                    // MILTIPUNTO
                    if (inventario_multipuntos.Es_MultiPunto(p.Value.cups22))
                    {
                        workSheet.Cells[f, c].Value = "Si";  
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        total_multipuntos++;
                    }
                    else
                    {
                        workSheet.Cells[f, c].Value = "No";  
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    c++;


                }

                foreach (KeyValuePair<string, EndesaEntity.facturacion.InventarioTipologias> p in bte.inventario)
                {
                                       

                    c = 1;
                    f++;

                    workSheet.Cells[f, c].Value = p.Value.cups13; c++;
                    workSheet.Cells[f, c].Value = p.Value.cups22; c++;
                    workSheet.Cells[f, c].Value = p.Value.cif; c++;
                    workSheet.Cells[f, c].Value = p.Value.razon_social; c++;

                    // Convertido
                    //workSheet.Cells[f, c].Value = "No";
                    //workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;


                    workSheet.Cells[f, c].Value = p.Value.empresa; c++;

                    if (!dic_agrup_empresa.TryGetValue(p.Value.empresa, out o))
                    {
                        o = new EndesaEntity.Agrupacion();
                        o.tipo = p.Value.empresa;
                        o.total = 1;
                        dic_agrup_empresa.Add(o.tipo, o);
                    }
                    else
                        o.total++;

                    workSheet.Cells[f, c].Value = p.Value.tarifa; c++;

                    if (!dic_agrup_tarifa.TryGetValue(p.Value.tarifa, out o))
                    {
                        o = new EndesaEntity.Agrupacion();
                        o.tipo = p.Value.tarifa;
                        o.total = 1;
                        dic_agrup_tarifa.Add(o.tipo, o);
                    }
                    else
                        o.total++;


                    // workSheet.Cells[f, c].Value = p.Value.tipo_punto_suministro; c++;
                    workSheet.Cells[f, c].Value = p.Value.provincia; c++;
                    workSheet.Cells[f, c].Value = p.Value.territorio; c++;
                    workSheet.Cells[f, c].Value = p.Value.contrato_ext; c++;
                    workSheet.Cells[f, c].Value = ""; c++; // AGRUPADA
                    // workSheet.Cells[f, c].Value = ""; c++; // AGE
                    workSheet.Cells[f, c].Value = ""; c++; // AGORA
                    workSheet.Cells[f, c].Value = ""; c++; // PASSTHOUGH	
                    workSheet.Cells[f, c].Value = ""; c++; // PASSPOOL HORARIO	
                    workSheet.Cells[f, c].Value = ""; c++; // PASSPOOL PERIODO	
                    workSheet.Cells[f, c].Value = ""; c++; // PASSPOOL SUBASTA	
                    workSheet.Cells[f, c].Value = ""; c++; // TIPO CONTRATO	
                    // workSheet.Cells[f, c].Value = ""; c++; // GESTIÓN PROPIA ATR	

                    valor_tam = tam.GetTAM(p.Value.cups22);

                    // TAM
                    if (valor_tam != -1111111)
                    {
                        workSheet.Cells[f, c].Value = valor_tam;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    }
                    else
                    {
                        workSheet.Cells[f, c].Value = "Sin Facturas";
                    }
                    c++;

                    workSheet.Cells[f, c].Value = p.Value.estado_contrato;  // ESTADO CONTRATO
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;
                    workSheet.Cells[f, c].Value = p.Value.version; c++; // VERSION

                    if (p.Value.fecha_puesta_servicio > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value.fecha_puesta_servicio;  // F. PUESTA SERVICIO
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;


                    workSheet.Cells[f, c].Value = "No";  // REVENDEDORES
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    //workSheet.Cells[f, c].Value = "No";  // ULTIMO CONVERTIDO
                    //workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //c++;

                    workSheet.Cells[f, c].Value = "No";  // MEDIDO EN BAJA
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    if (inventario_multipuntos.Es_MultiPunto(p.Value.cups22))
                    {
                        workSheet.Cells[f, c].Value = "Si";  // MEDIDO EN BAJA
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        total_multipuntos++;
                    }
                    else
                    {
                        workSheet.Cells[f, c].Value = "No";  // MEDIDO EN BAJA
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    c++;

                }

                foreach (KeyValuePair<string, EndesaEntity.facturacion.InventarioTipologias> p in mt.inventario)
                {
                    c = 1;
                    #region Cabecera
                    if (firstOnly)
                    {

                        workSheet.Cells[f, c].Value = "CUPS"; c++;
                        workSheet.Cells[f, c].Value = "CUPS22"; c++;
                        workSheet.Cells[f, c].Value = "CIF"; c++;
                        workSheet.Cells[f, c].Value = "RAZÓN SOCIAL"; c++;
                        workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                        workSheet.Cells[f, c].Value = "TARIFA"; c++;
                        workSheet.Cells[f, c].Value = "PROVINCIA"; c++;
                        workSheet.Cells[f, c].Value = "TERRITORIO"; c++;
                        workSheet.Cells[f, c].Value = "CONTRATO EXT"; c++;
                        workSheet.Cells[f, c].Value = "AGRUPADA"; c++;

                        workSheet.Cells[f, c].Value = "AGORA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                        
                        workSheet.Cells[f, c].Value = "PASSTHOUGH"; c++;
                        workSheet.Cells[f, c].Value = "PASSPOOL HORARIO"; c++;
                        workSheet.Cells[f, c].Value = "PASSPOOL PERIODO"; c++;
                        workSheet.Cells[f, c].Value = "PASSPOOL SUBASTA"; c++;
                        workSheet.Cells[f, c].Value = "TIPO CONTRATO"; c++;
                        // workSheet.Cells[f, c].Value = "GESTIÓN PROPIA ATR"; c++;
                        workSheet.Cells[f, c].Value = "TAM";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;
                        workSheet.Cells[f, c].Value = "ESTADO CONTRATO"; c++;
                        workSheet.Cells[f, c].Value = "VERSIÓN"; c++;
                        workSheet.Cells[f, c].Value = "F. PUESTA SERVICIO"; c++;
                        workSheet.Cells[f, c].Value = "REVENDEDORES"; c++;
                        // workSheet.Cells[f, c].Value = "ÚLTIMO CONVERTIDO"; c++;
                        workSheet.Cells[f, c].Value = "MEDIDO EN BAJA"; c++;
                        workSheet.Cells[f, c].Value = "MULTIPUNTO"; c++;

                        firstOnly = false;
                    }
                    #endregion

                    c = 1;
                    f++;

                    workSheet.Cells[f, c].Value = p.Value.cups13; c++;
                    workSheet.Cells[f, c].Value = p.Value.cups22; c++;
                    workSheet.Cells[f, c].Value = p.Value.cif; c++;
                    workSheet.Cells[f, c].Value = p.Value.razon_social; c++;

                    

                    workSheet.Cells[f, c].Value = p.Value.empresa; c++;

                    if (!dic_agrup_empresa.TryGetValue(p.Value.empresa, out o))
                    {
                        o = new EndesaEntity.Agrupacion();
                        o.tipo = p.Value.empresa;
                        o.total = 1;
                        dic_agrup_empresa.Add(o.tipo, o);
                    }
                    else
                        o.total++;

                    workSheet.Cells[f, c].Value = p.Value.tarifa; c++;

                    if (!dic_agrup_tarifa.TryGetValue(p.Value.tarifa, out o))
                    {
                        o = new EndesaEntity.Agrupacion();
                        o.tipo = p.Value.tarifa;
                        o.total = 1;
                        dic_agrup_tarifa.Add(o.tipo, o);
                    }
                    else
                        o.total++;


                    //workSheet.Cells[f, c].Value = p.Value.tipo_punto_suministro; c++;
                    workSheet.Cells[f, c].Value = p.Value.provincia; c++;
                    workSheet.Cells[f, c].Value = p.Value.territorio; c++;
                    workSheet.Cells[f, c].Value = p.Value.contrato_ext; c++;
                    workSheet.Cells[f, c].Value = ""; c++; // AGRUPADA
                    // workSheet.Cells[f, c].Value = ""; c++; // AGE
                    workSheet.Cells[f, c].Value = ""; c++; // AGORA
                    workSheet.Cells[f, c].Value = ""; c++; // PASSTHOUGH	
                    workSheet.Cells[f, c].Value = ""; c++; // PASSPOOL HORARIO	
                    workSheet.Cells[f, c].Value = ""; c++; // PASSPOOL PERIODO	
                    workSheet.Cells[f, c].Value = ""; c++; // PASSPOOL SUBASTA	
                    workSheet.Cells[f, c].Value = ""; c++; // TIPO CONTRATO	
                    //workSheet.Cells[f, c].Value = ""; c++; // GESTIÓN PROPIA ATR	

                    valor_tam = tam.GetTAM(p.Value.cups22);

                    // TAM
                    if (valor_tam != -1111111)
                    {
                        workSheet.Cells[f, c].Value = valor_tam;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    }
                    else
                    {
                        workSheet.Cells[f, c].Value = "Sin Facturas";
                    }
                    c++;



                    workSheet.Cells[f, c].Value = p.Value.estado_contrato;  // ESTADO CONTRATO
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;
                    workSheet.Cells[f, c].Value = p.Value.version; c++; // VERSION

                    if(p.Value.fecha_puesta_servicio > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value.fecha_puesta_servicio;  // F. PUESTA SERVICIO
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; 
                    }
                    c++;


                    workSheet.Cells[f, c].Value = "No";  // REVENDEDORES
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    //workSheet.Cells[f, c].Value = "No";  // ULTIMO CONVERTIDO
                    //workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //c++;

                    workSheet.Cells[f, c].Value = "No";  // MEDIDO EN BAJA
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    c++;

                    if (inventario_multipuntos.Es_MultiPunto(p.Value.cups22))
                    {
                        workSheet.Cells[f, c].Value = "Si";  // MEDIDO EN BAJA
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        total_multipuntos++;
                    }
                    else
                    {
                        workSheet.Cells[f, c].Value = "No";  // MEDIDO EN BAJA
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                   
                    c++;

                }


                #region guardado de totales
                foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_empresa)
                    GuardaTotales(DateTime.Now, "EMPRESA", p.Key, p.Value.total);

                foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_tarifa)
                    GuardaTotales(DateTime.Now, "TARIFA", p.Key, p.Value.total);

                GuardaTotales(DateTime.Now, "TOTALES", "TOTAL AGRUPADAS", total_agrupada);
                GuardaTotales(DateTime.Now, "TOTALES", "TOTAL AGE", total_age);
                GuardaTotales(DateTime.Now, "TOTALES", "TOTAL AGORA", total_agora);
                GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PASSTHOUGH", total_passthough);
                GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PASSPOOL HORARIO", total_passpool_horario);
                GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PASSPOOL SUBASTA", total_passpool_subasta);
                GuardaTotales(DateTime.Now, "TOTALES", "TOTAL GESTIÓN PROPIA ATR", total_gestion_atr);
                GuardaTotales(DateTime.Now, "TOTALES", "TOTAL REVENDEDORES", total_revendedores);
                GuardaTotales(DateTime.Now, "TOTALES", "TOTAL MEDIDO EN BAJA", total_medido_baja);
                GuardaTotales(DateTime.Now, "TOTALES", "TOTAL MULTIPUNTO", total_multipuntos);


                dic_totales = CargaTotales();

                #endregion


                var allCells = workSheet.Cells[1, 1, f, c];
                headerCells = workSheet.Cells[1, 1, 1, c];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;
                allCells = workSheet.Cells[1, 1, f, c];

                allCells.AutoFitColumns();

                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:W1"].AutoFilter = true;
                allCells.AutoFitColumns();





                workSheet = excelPackage.Workbook.Worksheets.Add("AGRUPACIONES");
                headerFont.Bold = true;
                f = 1;
                c = 1;


                workSheet.Cells[f, c].Value = "EMPRESA";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                c++;
                     

                foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_empresa)
                {
                    c = 1;
                    f++;
                    workSheet.Cells[f, c].Value = p.Key; c++;                 
                    c++;
                }

                f++;
                f++;
                c = 1;
                workSheet.Cells[f, c].Value = "TARIFA";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                c++;

                foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_tarifa)
                {
                    c = 1;
                    f++;
                    workSheet.Cells[f, c].Value = p.Key; c++;
                    //workSheet.Cells[f, c].Value = p.Value.total;
                    //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";


                }
                f++;
                f++;
                c = 1;

                #region Cabecera lateral Totales
                //foreach(string p in GetListaConceptos("TOTALES"))
                //{
                //    workSheet.Cells[f, c].Value = p + ":";
                //    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //    f++;
                //    c = 1;
                //}
                #endregion

                workSheet.Cells[f, c].Value = "TOTAL AGRUPADAS:";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                c++;
                //workSheet.Cells[f, c].Value = total_agrupada;
                //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                //f++;
                c = 1;

                //workSheet.Cells[f, c].Value = "TOTAL AGE:";
                //workSheet.Cells[f, c].Style.Font.Bold = true;
                //workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //c++;
                //workSheet.Cells[f, c].Value = total_age;
                //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                f++;
                c = 1;

                workSheet.Cells[f, c].Value = "TOTAL AGORA:";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                c++;
                //workSheet.Cells[f, c].Value = total_agora;
                //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                f++;
                c = 1;

                workSheet.Cells[f, c].Value = "TOTAL PASSTHOUGH:";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                c++;
                //workSheet.Cells[f, c].Value = total_passthough;
                //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                f++;
                c = 1;

                workSheet.Cells[f, c].Value = "TOTAL PASSPOOL HORARIO:";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                c++;
                //workSheet.Cells[f, c].Value = total_passpool_horario;
                //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                f++;
                c = 1;

                workSheet.Cells[f, c].Value = "TOTAL PASSPOOL PERIODO:";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                c++;
                //workSheet.Cells[f, c].Value = total_passpool_periodo;
                //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                //GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PASSPOOL PERIODO", total_passpool_periodo);
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = "TOTAL PASSPOOL SUBASTA:";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                c++;
                //workSheet.Cells[f, c].Value = total_passpool_subasta;
                //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";                
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = "TOTAL GESTIÓN PROPIA ATR:";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                c++;
                //workSheet.Cells[f, c].Value = total_gestion_atr;
                //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";                
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = "TOTAL REVENDEDORES:";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                c++;
                //workSheet.Cells[f, c].Value = total_revendedores;
                //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";                
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = "TOTAL MEDIDO EN BAJA:";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                c++;
                //workSheet.Cells[f, c].Value = total_medido_baja;
                //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";                
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = "TOTAL MULTIPUNTO:";
                workSheet.Cells[f, c].Style.Font.Bold = true;
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                firstOnly = true;
                foreach (KeyValuePair<DateTime, List<EndesaEntity.facturacion.IRF_Totales_Hist>> p in dic_totales)
                {
                    f = 1;

                    #region difenrecia
                    if (firstOnly)
                    {
                        c++;
                        firstOnly = false;
                        workSheet.Cells[f, c].Value = "DIF (D-C)";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        f = 2;
                        workSheet.Cells[f, c].Value =
                            GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "EMPRESA", workSheet.Cells[f, 1].Value.ToString()) -
                            GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f = 3;
                        workSheet.Cells[f, c].Value =
                            GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "EMPRESA", workSheet.Cells[f, 1].Value.ToString()) -
                            GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f = 4;
                        workSheet.Cells[f, c].Value =
                            GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "EMPRESA", workSheet.Cells[f, 1].Value.ToString()) -
                            GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        f++;
                        foreach (KeyValuePair<string, EndesaEntity.Agrupacion> tarifa in dic_agrup_tarifa)
                        {

                            f++;
                            workSheet.Cells[f, c].Value =
                                GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TARIFA", tarifa.Key) -
                                GetTotal(p.Key, "TARIFA", tarifa.Key);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        }


                        f++;
                        f++;
                        workSheet.Cells[f, c].Value =
                            GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL AGRUPADAS") -
                            GetTotal(p.Key, "TOTALES", "TOTAL AGRUPADAS");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        //f++;
                        //workSheet.Cells[f, c].Value =
                        //    GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL AGE") -
                        //    GetTotal(p.Key, "TOTALES", "TOTAL AGE");
                        //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value =
                            GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL AGORA") -
                            GetTotal(p.Key, "TOTALES", "TOTAL AGORA");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value =
                            GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL PASSTHOUGH") -
                            GetTotal(p.Key, "TOTALES", "TOTAL PASSTHOUGH");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value =
                            GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL PASSPOOL HORARIO") -
                            GetTotal(p.Key, "TOTALES", "TOTAL PASSPOOL HORARIO");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value =
                            GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL PASSPOOL PERIODO") -
                            GetTotal(p.Key, "TOTALES", "TOTAL PASSPOOL PERIODO");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value =
                            GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL PASSPOOL SUBASTA") -
                            GetTotal(p.Key, "TOTALES", "TOTAL PASSPOOL SUBASTA");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value =
                            GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL GESTIÓN PROPIA ATR") -
                            GetTotal(p.Key, "TOTALES", "TOTAL GESTIÓN PROPIA ATR");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value =
                            GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL REVENDEDORES") -
                            GetTotal(p.Key, "TOTALES", "TOTAL REVENDEDORES");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value =
                            GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL MEDIDO EN BAJA") -
                            GetTotal(p.Key, "TOTALES", "TOTAL MEDIDO EN BAJA");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        f++;
                        workSheet.Cells[f, c].Value =
                            GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL MULTIPUNTO") -
                            GetTotal(p.Key, "TOTALES", "TOTAL MULTIPUNTO");
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    }

                    #endregion

                    f = 1;

                    c++;
                    workSheet.Cells[f, c].Value = p.Key;
                    workSheet.Cells[f, c].Style.Font.Bold = true;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                    f++;
                    workSheet.Cells[f, c].Value = GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    f++;
                    workSheet.Cells[f, c].Value = GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    f++;
                    workSheet.Cells[f, c].Value = GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    f++;
                    f++;
                    foreach (KeyValuePair<string, EndesaEntity.Agrupacion> tarifa in dic_agrup_tarifa)
                    {

                        f++;
                        workSheet.Cells[f, c].Value = GetTotal(p.Key, "TARIFA", tarifa.Key);
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }

                    f++;
                    f++;
                    workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL AGRUPADAS");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    //f++;
                    //workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL AGE");
                    //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    f++;
                    workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL AGORA");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    f++;
                    workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL PASSTHOUGH");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    f++;
                    workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL PASSPOOL HORARIO");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    f++;
                    workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL PASSPOOL PERIODO");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    f++;
                    workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL PASSPOOL SUBASTA");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    f++;
                    workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL GESTIÓN PROPIA ATR");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    f++;
                    workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL REVENDEDORES");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    f++;
                    workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL MEDIDO EN BAJA");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                    f++;
                    workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL MULTIPUNTO");
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                }



                allCells = workSheet.Cells[1, 1, f, c];
                allCells.AutoFitColumns();


                excelPackage.Save();

                if (param.GetValue("EnviarMail") == "S")
                    EnvioCorreoInformeTipologias(ruta_salida_archivo);


            }
            catch(Exception ex)
            {
                ficheroLog.addError("Inventario_por_Tipologias: " + ex.Message);
            }

            
        }

        public void Inventario_por_Tipologias_v2(string fichero)
        {
            EndesaBusiness.contratacion.ContratosPS_Portugal_BTN btn = new contratacion.ContratosPS_Portugal_BTN();
            EndesaBusiness.contratacion.ContratosPS_Portugal_BTE bte = new contratacion.ContratosPS_Portugal_BTE();
            EndesaBusiness.contratacion.ContratosPS_Portugal_MT mt = new contratacion.ContratosPS_Portugal_MT();

            bool firstOnly = true;
            int f = 1;
            int c = 1;

            EndesaEntity.Agrupacion o;


            // Totales para agrupaciones

            int total_agrupada = 0;            
            int total_agora = 0;            
            int total_multipuntos = 0;
            int total_pass_pool_58 = 0;
            int total_pass_pool_59 = 0;
            int total_gestion_propia_atr = 0;
            int total_iva_normal = 0;
            int total_iva_intermedio = 0;
            int total_iva_reducido = 0;



            double valor_tam = 0;
            int total_reg = 0;

            string ruta_salida_archivo;

            bool mostrar_btn = true;
            bool mostrar_bte = true;
            bool mostrar_mt = true;
            bool mostrar_agrupaciones = true;

            DirectoryInfo dirSalida;
            FileInfo file;

            try
            {
                if (ss_pp.GetFecha_FinProceso("Facturación", "Informe Inventario Tipologias PT", "Informe Inventario Tipologias PT").Date 
                    < ss_pp.GetFecha_FinProceso("Facturación", "Informe Pendiente BI", "Informe Pendiente BI").Date)
                {
                    ss_pp.Update_Fecha_Inicio("Facturación", "Informe Inventario Tipologias PT", "Informe Inventario Tipologias PT");
                    
                    string[] listaArchivos = System.IO.Directory.GetFiles(param.GetValue("Ubicacion_Informes"),
                   param.GetValue("Excel_prefijo_Tipologias") + "*.xlsx");

                    for (int i = 0; i < listaArchivos.Length; i++)
                    {
                        file = new FileInfo(listaArchivos[i]);
                        file.Delete();
                    }


                    // Carga MultiPuntos COMPOR
                    EndesaBusiness.compor.Inventario_Multipuntos inventario_multipuntos
                        = new compor.Inventario_Multipuntos();

                    // Carga Exenciones
                    dic_exenciones = CargaExenciones();

                    EndesaBusiness.medida.TAM tam = new medida.TAM();
                    tam.CargaTAM_NIF_CUPS();


                    // Carga Agora Manual
                    // Se buscan los contratos vigentes a la hora de 
                    // sacar el informe.
                    EndesaBusiness.facturacion.Agora_DyP agoraManual = new Agora_DyP();

                    // Agrupadas
                    EndesaBusiness.contratacion.Agrupadas agrupadas = new contratacion.Agrupadas();

                    //tam = new medida.TAM();
                    //tam.CargaTAM();





                    if (fichero != null)
                        ruta_salida_archivo = fichero;
                    else
                    {
                        ruta_salida_archivo = param.GetValue("Ubicacion_Informes")
                            + param.GetValue("Excel_prefijo_Tipologias") + "_"
                            + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
                    }


                    FileInfo nombreSalidaExcel = new FileInfo(ruta_salida_archivo);


                    // Ruta de la plantilla 
                    FileInfo plantillaExcel = new FileInfo(System.Environment.CurrentDirectory + param.GetValue("PlantillaInforme"));
                    //dirSalida = new DirectoryInfo(param.GetValue("SalidaInforme"));



                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                    ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);
                    var workSheet = excelPackage.Workbook.Worksheets["PS"];


                    utilidades.Fechas utilfecha = new Fechas();


                    #region Cabecera
                    c = 1;


                    //workSheet.Cells[f, c].Value = "CUPS"; c++;
                    //workSheet.Cells[f, c].Value = "CUPS20"; c++;
                    //workSheet.Cells[f, c].Value = "CIF"; c++;
                    //workSheet.Cells[f, c].Value = "RAZÓN SOCIAL"; c++;
                    //workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                    //workSheet.Cells[f, c].Value = "PROVINCIA"; c++;
                    //workSheet.Cells[f, c].Value = "TERRITORIO"; c++;
                    //workSheet.Cells[f, c].Value = "CONTRATO EXT"; c++;
                    //workSheet.Cells[f, c].Value = "TIPO PUNTO PRODUCTO"; c++;
                    //workSheet.Cells[f, c].Value = "PRODUCTO"; c++;
                    //workSheet.Cells[f, c].Value = "TARIFA"; c++;
                    //workSheet.Cells[f, c].Value = "TIPO TARIFA"; c++;
                    //workSheet.Cells[f, c].Value = "MUILTIPUNTO"; c++;

                    //workSheet.Cells[f, c].Value = "CICLO P01 MT";
                    //workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                    //workSheet.Cells[f, c].Value = "GESTIÓN PROPIA ATR"; c++;
                    //workSheet.Cells[f, c].Value = "AGRUPADA"; c++;
                    //workSheet.Cells[f, c].Value = "TIPO EXENCION"; c++;
                    //workSheet.Cells[f, c].Value = "IVA NORMAL"; c++;
                    //workSheet.Cells[f, c].Value = "IVA INTERMEDIO BTN"; c++;
                    //workSheet.Cells[f, c].Value = "IVA REDUCIDO BTN"; c++;
                    //workSheet.Cells[f, c].Value = "PASS POOL 58"; c++;
                    //workSheet.Cells[f, c].Value = "PASS POOL 59"; c++;
                    //workSheet.Cells[f, c].Value = "AGORA";
                    //workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                    #endregion

                    EndesaBusiness.compor.Inventario inventario =
                            new compor.Inventario();

                    if (mostrar_btn)
                    {
                        btn.CargaInventarioBTN();

                        foreach (KeyValuePair<string, EndesaEntity.facturacion.InventarioTipologias> p in btn.inventario)
                        {
                            c = 1;
                            f++;

                            inventario.Get_Info_Inventario(p.Value.cups22);

                            workSheet.Cells[f, c].Value = p.Value.cups13; c++;
                            workSheet.Cells[f, c].Value = p.Value.cups22; c++;

                            workSheet.Cells[f, c].Value = p.Value.estado_contrato;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            if (p.Value.fecha_alta > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.Value.fecha_alta;  // F. alta
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;

                            if (p.Value.fecha_baja > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.Value.fecha_baja;  // F. Baja
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;


                            workSheet.Cells[f, c].Value = tam.GetTAM(p.Value.cif, p.Value.cups22.Substring(0, 20));
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";

                            c++;

                            workSheet.Cells[f, c].Value = p.Value.cif; c++;
                            workSheet.Cells[f, c].Value = p.Value.razon_social; c++;
                            workSheet.Cells[f, c].Value = p.Value.empresa; c++;
                            workSheet.Cells[f, c].Value = p.Value.provincia; c++;
                            workSheet.Cells[f, c].Value = p.Value.territorio; c++;
                            workSheet.Cells[f, c].Value = p.Value.contrato_ext; c++;

                            // TIPO PUNTO PRUDUCTO
                            if (inventario.existe)
                                workSheet.Cells[f, c].Value = inventario.tipo;

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // PRODUCTO                        
                            workSheet.Cells[f, c].Value = btn.GetProducto(p.Value.cups22);
                            c++;

                            workSheet.Cells[f, c].Value = p.Value.tarifa; c++;
                            workSheet.Cells[f, c].Value = p.Value.tipo_tarifa; c++;

                            if (inventario_multipuntos.Es_MultiPunto(p.Value.cups22))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_multipuntos++;
                            }
                            workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // Ciclo_P01_MT o CALENDARIO

                            if (inventario.existe)
                                workSheet.Cells[f, c].Value = inventario.calendario; c++;

                            // GESTION PROPIA ATR
                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;


                            // AGRUPADA
                            if (agrupadas.Agrupada(p.Value.contrato_ext))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_agrupada++;
                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // TIPO EXENCION
                            workSheet.Cells[f, c].Value = Get_TipoAjuste(p.Value.cups22); c++;

                            // IVA NORMAL

                            if (inventario.existe)
                            {
                                if (p.Value.kaveas > 6.9)
                                {
                                    workSheet.Cells[f, c].Value = "Sí";
                                    total_iva_normal++;
                                }

                                else
                                    workSheet.Cells[f, c].Value = "No";

                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            // IVA INTERMEDIO BTN

                            if (inventario.existe)
                            {
                                if (inventario.tipo != "ILUMINACIÓN PÚBLICA" && p.Value.kaveas <= 6.9)
                                {
                                    workSheet.Cells[f, c].Value = "Sí";
                                    total_iva_intermedio++;
                                }
                                else
                                    workSheet.Cells[f, c].Value = "No";

                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            // IVA REDUCIDO BTN
                            if (p.Value.kaveas <= 3.45)
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_iva_reducido++;
                            }

                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            // IVA reducido CAV (6%)
                            workSheet.Cells[f, c].Value = "Sí";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            // PASSPOOL 58

                            if (btn.TieneComplemento(p.Value.cups22, "P58"))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_pass_pool_58++;
                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // PASSPOOL 59
                            if (btn.TieneComplemento(p.Value.cups22, "P59"))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_pass_pool_59++;
                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // AGORA
                            if (agoraManual.EsAgoraManual(p.Value.cups22))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_agora++;
                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // PERIODO PENDIENTE
                            workSheet.Cells[f, c].Value = p.Value.periodo_pendiente;
                            c++;
                            // ESTADO PENDIENTE
                            workSheet.Cells[f, c].Value = p.Value.estado_pendiente;
                            c++;
                            // ESTADO PENDIENTE
                            workSheet.Cells[f, c].Value = p.Value.subestado_pendiente;
                            c++;

                            #region TOTALES
                            // EMPRESA
                            if (!dic_agrup_empresa.TryGetValue(p.Value.empresa, out o))
                            {
                                o = new EndesaEntity.Agrupacion();
                                o.tipo = p.Value.empresa;
                                o.total = 1;
                                dic_agrup_empresa.Add(o.tipo, o);
                            }
                            else
                                o.total++;

                            //TARIFA
                            if (!dic_agrup_tarifa.TryGetValue(p.Value.tarifa, out o))
                            {
                                o = new EndesaEntity.Agrupacion();
                                o.tipo = p.Value.tarifa;
                                o.total = 1;
                                dic_agrup_tarifa.Add(o.tipo, o);
                            }
                            else
                                o.total++;

                            //TIPO TARIFA
                            if (!dic_agrup_tipo_tarifa.TryGetValue(p.Value.tarifa, out o))
                            {
                                o = new EndesaEntity.Agrupacion();
                                o.tipo = p.Value.tarifa;
                                o.total = 1;
                                dic_agrup_tipo_tarifa.Add(o.tipo, o);
                            }
                            else
                                o.total++;


                            //PRODUCTO
                            if (!dic_agrup_producto.TryGetValue(btn.GetProducto(p.Value.cups22), out o))
                            {
                                o = new EndesaEntity.Agrupacion();
                                o.tipo = btn.GetProducto(p.Value.cups22);
                                o.total = 1;
                                dic_agrup_producto.Add(o.tipo, o);
                            }
                            else
                                o.total++;

                            // TIPO PRODUCTO
                            if (inventario.existe)
                            {
                                if (!dic_agrup_tipo_punto_producto.TryGetValue(inventario.tipo, out o))
                                {
                                    o = new EndesaEntity.Agrupacion();
                                    o.tipo = inventario.tipo;
                                    o.total = 1;
                                    dic_agrup_tipo_punto_producto.Add(o.tipo, o);
                                }
                                else
                                    o.total++;
                            }

                            if (inventario.existe)
                            {
                                if (!dic_agrup_ciclo_p01_mt.TryGetValue(inventario.calendario, out o))
                                {
                                    o = new EndesaEntity.Agrupacion();
                                    o.tipo = inventario.calendario;
                                    o.total = 1;
                                    dic_agrup_ciclo_p01_mt.Add(o.tipo, o);
                                }
                                else
                                    o.total++;
                            }


                            if (Get_TipoAjuste(p.Value.cups22) != "")
                            {
                                if (!dic_agrup_exencion.TryGetValue(Get_TipoAjuste(p.Value.cups22), out o))
                                {
                                    o = new EndesaEntity.Agrupacion();
                                    o.tipo = Get_TipoAjuste(p.Value.cups22);
                                    o.total = 1;
                                    dic_agrup_exencion.Add(o.tipo, o);
                                }
                                else
                                    o.total++;
                            }



                            #endregion



                        }
                    }

                    if (mostrar_bte)
                    {
                        bte.CargaInventarioBTE();


                        foreach (KeyValuePair<string, EndesaEntity.facturacion.InventarioTipologias> p in bte.inventario)
                        {

                            c = 1;
                            f++;

                            inventario.Get_Info_Inventario(p.Value.cups22);

                            workSheet.Cells[f, c].Value = p.Value.cups13; c++;
                            workSheet.Cells[f, c].Value = p.Value.cups22; c++;

                            workSheet.Cells[f, c].Value = p.Value.estado_contrato;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            if (p.Value.fecha_alta > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.Value.fecha_alta;  // F. alta
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;

                            if (p.Value.fecha_baja > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.Value.fecha_baja;  // F. Baja
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;

                            workSheet.Cells[f, c].Value = tam.GetTAM(p.Value.cif, p.Value.cups22.Substring(0, 20));
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;


                            workSheet.Cells[f, c].Value = p.Value.cif; c++;
                            workSheet.Cells[f, c].Value = p.Value.razon_social; c++;
                            workSheet.Cells[f, c].Value = p.Value.empresa; c++;
                            workSheet.Cells[f, c].Value = p.Value.provincia; c++;
                            workSheet.Cells[f, c].Value = p.Value.territorio; c++;
                            workSheet.Cells[f, c].Value = p.Value.contrato_ext; c++;

                            // TIPO PUNTO PRUDUCTO
                            if (inventario.existe)
                                workSheet.Cells[f, c].Value = inventario.tipo;

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // PRODUCTO                        
                            workSheet.Cells[f, c].Value = bte.GetProducto(p.Value.cups22);
                            c++;

                            workSheet.Cells[f, c].Value = p.Value.tarifa; c++;
                            workSheet.Cells[f, c].Value = p.Value.tipo_tarifa; c++;

                            if (inventario_multipuntos.Es_MultiPunto(p.Value.cups22))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_multipuntos++;
                            }
                            workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // Ciclo_P01_MT o CALENDARIO

                            if (inventario.existe)
                                workSheet.Cells[f, c].Value = inventario.calendario; c++;

                            // GESTION PROPIA ATR
                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;


                            // AGRUPADA
                            if (agrupadas.Agrupada(p.Value.contrato_ext))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_agrupada++;
                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // TIPO EXENCION
                            workSheet.Cells[f, c].Value = Get_TipoAjuste(p.Value.cups22); c++;

                            // IVA NORMAL


                            workSheet.Cells[f, c].Value = "Sí";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            // IVA INTERMEDIO BTN


                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            // IVA REDUCIDO BTN

                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            // IVA reducido CAV (6%)
                            workSheet.Cells[f, c].Value = "Sí";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;


                            // PASSPOOL 58

                            if (bte.TieneComplemento(p.Value.cups22, "P58"))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_pass_pool_58++;
                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // PASSPOOL 59
                            if (bte.TieneComplemento(p.Value.cups22, "P59"))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_pass_pool_59++;
                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // AGORA
                            if (agoraManual.EsAgoraManual(p.Value.cups22))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_agora++;
                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // PERIODO PENDIENTE
                            workSheet.Cells[f, c].Value = p.Value.periodo_pendiente;
                            c++;
                            // ESTADO PENDIENTE
                            workSheet.Cells[f, c].Value = p.Value.estado_pendiente;
                            c++;
                            // ESTADO PENDIENTE
                            workSheet.Cells[f, c].Value = p.Value.subestado_pendiente;
                            c++;

                            #region TOTALES
                            // EMPRESA
                            if (!dic_agrup_empresa.TryGetValue(p.Value.empresa, out o))
                            {
                                o = new EndesaEntity.Agrupacion();
                                o.tipo = p.Value.empresa;
                                o.total = 1;
                                dic_agrup_empresa.Add(o.tipo, o);
                            }
                            else
                                o.total++;

                            //TARIFA
                            if (!dic_agrup_tarifa.TryGetValue(p.Value.tarifa, out o))
                            {
                                o = new EndesaEntity.Agrupacion();
                                o.tipo = p.Value.tarifa;
                                o.total = 1;
                                dic_agrup_tarifa.Add(o.tipo, o);
                            }
                            else
                                o.total++;

                            //PRODUCTO
                            if (!dic_agrup_producto.TryGetValue(bte.GetProducto(p.Value.cups22), out o))
                            {
                                o = new EndesaEntity.Agrupacion();
                                o.tipo = bte.GetProducto(p.Value.cups22);
                                o.total = 1;
                                dic_agrup_producto.Add(o.tipo, o);
                            }
                            else
                                o.total++;

                            // TIPO PRODUCTO
                            if (inventario.existe)
                            {
                                if (!dic_agrup_tipo_punto_producto.TryGetValue(inventario.tipo, out o))
                                {
                                    o = new EndesaEntity.Agrupacion();
                                    o.tipo = inventario.tipo;
                                    o.total = 1;
                                    dic_agrup_tipo_punto_producto.Add(o.tipo, o);
                                }
                                else
                                    o.total++;
                            }

                            if (inventario.existe)
                            {
                                if (!dic_agrup_ciclo_p01_mt.TryGetValue(inventario.calendario, out o))
                                {
                                    o = new EndesaEntity.Agrupacion();
                                    o.tipo = inventario.calendario;
                                    o.total = 1;
                                    dic_agrup_ciclo_p01_mt.Add(o.tipo, o);
                                }
                                else
                                    o.total++;
                            }


                            if (Get_TipoAjuste(p.Value.cups22) != "")
                            {
                                if (!dic_agrup_exencion.TryGetValue(Get_TipoAjuste(p.Value.cups22), out o))
                                {
                                    o = new EndesaEntity.Agrupacion();
                                    o.tipo = Get_TipoAjuste(p.Value.cups22);
                                    o.total = 1;
                                    dic_agrup_exencion.Add(o.tipo, o);
                                }
                                else
                                    o.total++;
                            }



                            #endregion


                        }
                    }

                    if (mostrar_mt)
                    {

                        mt.CargaInventarioMT();

                        foreach (KeyValuePair<string, EndesaEntity.facturacion.InventarioTipologias> p in mt.inventario)
                        {

                            c = 1;
                            f++;

                            inventario.Get_Info_Inventario(p.Value.cups22);

                            workSheet.Cells[f, c].Value = p.Value.cups13; c++;
                            workSheet.Cells[f, c].Value = p.Value.cups22; c++;

                            workSheet.Cells[f, c].Value = p.Value.estado_contrato;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            if (p.Value.fecha_alta > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.Value.fecha_alta;  // F. alta
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;

                            if (p.Value.fecha_baja > DateTime.MinValue)
                            {
                                workSheet.Cells[f, c].Value = p.Value.fecha_baja;  // F. Baja
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            c++;

                            workSheet.Cells[f, c].Value = tam.GetTAM(p.Value.cif, p.Value.cups22.Substring(0, 20));
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;

                            workSheet.Cells[f, c].Value = p.Value.cif; c++;
                            workSheet.Cells[f, c].Value = p.Value.razon_social; c++;
                            workSheet.Cells[f, c].Value = p.Value.empresa; c++;
                            workSheet.Cells[f, c].Value = p.Value.provincia; c++;
                            workSheet.Cells[f, c].Value = p.Value.territorio; c++;
                            workSheet.Cells[f, c].Value = p.Value.contrato_ext; c++;

                            // TIPO PUNTO PRUDUCTO
                            if (inventario.existe)
                                workSheet.Cells[f, c].Value = inventario.tipo;

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // PRODUCTO                        
                            workSheet.Cells[f, c].Value = mt.GetProducto(p.Value.cups22);
                            c++;

                            workSheet.Cells[f, c].Value = p.Value.tarifa; c++;
                            workSheet.Cells[f, c].Value = p.Value.tipo_tarifa; c++;

                            if (inventario_multipuntos.Es_MultiPunto(p.Value.cups22))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_multipuntos++;
                            }
                            workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // Ciclo_P01_MT o CALENDARIO

                            if (inventario.existe)
                                workSheet.Cells[f, c].Value = inventario.calendario; c++;

                            // GESTION PROPIA ATR
                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;


                            // AGRUPADA
                            if (agrupadas.Agrupada(p.Value.contrato_ext))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_agrupada++;
                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // TIPO EXENCION
                            workSheet.Cells[f, c].Value = Get_TipoAjuste(p.Value.cups22); c++;

                            // IVA NORMAL


                            workSheet.Cells[f, c].Value = "Sí";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            // IVA INTERMEDIO BTN


                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            // IVA REDUCIDO BTN

                            workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            // IVA reducido CAV (6%)
                            workSheet.Cells[f, c].Value = "Sí";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                            // PASSPOOL 58

                            if (mt.TieneComplemento(p.Value.cups22, "P58"))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_pass_pool_58++;
                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // PASSPOOL 59
                            if (mt.TieneComplemento(p.Value.cups22, "P59"))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_pass_pool_59++;
                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // AGORA
                            if (agoraManual.EsAgoraManual(p.Value.cups22))
                            {
                                workSheet.Cells[f, c].Value = "Sí";
                                total_agora++;
                            }
                            else
                                workSheet.Cells[f, c].Value = "No";

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            c++;

                            // PERIODO PENDIENTE
                            workSheet.Cells[f, c].Value = p.Value.periodo_pendiente;
                            c++;
                            // ESTADO PENDIENTE
                            workSheet.Cells[f, c].Value = p.Value.estado_pendiente;
                            c++;
                            // ESTADO PENDIENTE
                            workSheet.Cells[f, c].Value = p.Value.subestado_pendiente;
                            c++;

                            #region TOTALES
                            // EMPRESA
                            if (!dic_agrup_empresa.TryGetValue(p.Value.empresa, out o))
                            {
                                o = new EndesaEntity.Agrupacion();
                                o.tipo = p.Value.empresa;
                                o.total = 1;
                                dic_agrup_empresa.Add(o.tipo, o);
                            }
                            else
                                o.total++;

                            //TARIFA
                            if (!dic_agrup_tarifa.TryGetValue(p.Value.tarifa, out o))
                            {
                                o = new EndesaEntity.Agrupacion();
                                o.tipo = p.Value.tarifa;
                                o.total = 1;
                                dic_agrup_tarifa.Add(o.tipo, o);
                            }
                            else
                                o.total++;

                            //PRODUCTO
                            if (!dic_agrup_producto.TryGetValue(mt.GetProducto(p.Value.cups22), out o))
                            {
                                o = new EndesaEntity.Agrupacion();
                                o.tipo = mt.GetProducto(p.Value.cups22);
                                o.total = 1;
                                dic_agrup_producto.Add(o.tipo, o);
                            }
                            else
                                o.total++;

                            // TIPO PRODUCTO
                            if (inventario.existe)
                            {
                                if (!dic_agrup_tipo_punto_producto.TryGetValue(inventario.tipo, out o))
                                {
                                    o = new EndesaEntity.Agrupacion();
                                    o.tipo = inventario.tipo;
                                    o.total = 1;
                                    dic_agrup_tipo_punto_producto.Add(o.tipo, o);
                                }
                                else
                                    o.total++;
                            }

                            if (inventario.existe)
                            {
                                if (!dic_agrup_ciclo_p01_mt.TryGetValue(inventario.calendario, out o))
                                {
                                    o = new EndesaEntity.Agrupacion();
                                    o.tipo = inventario.calendario;
                                    o.total = 1;
                                    dic_agrup_ciclo_p01_mt.Add(o.tipo, o);
                                }
                                else
                                    o.total++;
                            }


                            if (Get_TipoAjuste(p.Value.cups22) != "")
                            {
                                if (!dic_agrup_exencion.TryGetValue(Get_TipoAjuste(p.Value.cups22), out o))
                                {
                                    o = new EndesaEntity.Agrupacion();
                                    o.tipo = Get_TipoAjuste(p.Value.cups22);
                                    o.total = 1;
                                    dic_agrup_exencion.Add(o.tipo, o);
                                }
                                else
                                    o.total++;
                            }



                            #endregion


                        }
                    }



                    #region guardado de totales
                    foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_empresa)
                        GuardaTotales(DateTime.Now, "EMPRESA", p.Key, p.Value.total);

                    foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_tarifa)
                        GuardaTotales(DateTime.Now, "TARIFA", p.Key, p.Value.total);

                    foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_producto)
                        GuardaTotales(DateTime.Now, "PRODUCTO", p.Key, p.Value.total);

                    foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_tipo_punto_producto)
                        GuardaTotales(DateTime.Now, "TIPO PUNTO PRODUCTO", p.Key, p.Value.total);

                    foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_tipo_tarifa)
                        GuardaTotales(DateTime.Now, "TIPO TARIFA", p.Key, p.Value.total);

                    foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_ciclo_p01_mt)
                        GuardaTotales(DateTime.Now, "TIPO CICLO", p.Key, p.Value.total);

                    foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_exencion)
                        GuardaTotales(DateTime.Now, "TIPO EXENCION", p.Key, p.Value.total);



                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL GESTION PROPIA ATR", total_gestion_propia_atr);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PASS POOL 58", total_pass_pool_58);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PASS POOL 59", total_pass_pool_59);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL IVA NORMAL", total_iva_normal);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL IVA INTERMEDIO", total_iva_intermedio);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL IVA REDUCIDO", total_iva_reducido);

                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL AGRUPADAS", total_agrupada);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL AGORA", total_agora);
                    GuardaTotales(DateTime.Now, "TOTALES", "TOTAL MULTIPUNTO", total_multipuntos);




                    dic_totales = CargaTotales();

                    #endregion


                    var allCells = workSheet.Cells[1, 1, f, c];
                    var headerCells = workSheet.Cells[1, 1, 1, c];
                    var headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;
                    allCells = workSheet.Cells[1, 1, f, c];

                    allCells.AutoFitColumns();

                    workSheet.View.FreezePanes(2, 1);
                    workSheet.Cells["A1:X1"].AutoFilter = true;
                    allCells.AutoFitColumns();




                    if (mostrar_agrupaciones)
                    {

                        workSheet = excelPackage.Workbook.Worksheets["AGRUPACIONES"];

                        headerFont.Bold = true;
                        f = 1;
                        c = 1;


                        workSheet.Cells[f, c].Value = "EMPRESA";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;


                        foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_empresa)
                        {
                            c = 1;
                            f++;
                            workSheet.Cells[f, c].Value = p.Key; c++;
                            c++;
                        }

                        f++;
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = "TARIFA";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_tarifa)
                        {
                            c = 1;
                            f++;
                            workSheet.Cells[f, c].Value = p.Key;
                            c++;
                        }


                        f++;
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = "TIPO PUNTO PRODUCTO";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_tipo_punto_producto)
                        {
                            c = 1;
                            f++;
                            workSheet.Cells[f, c].Value = p.Key;
                            c++;

                        }

                        f++;
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = "TIPO TARIFA";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_tipo_tarifa)
                        {
                            c = 1;
                            f++;
                            workSheet.Cells[f, c].Value = p.Key;
                            c++;

                        }

                        f++;
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = "EXENCIONES";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;

                        foreach (KeyValuePair<string, EndesaEntity.Agrupacion> p in dic_agrup_exencion)
                        {
                            c = 1;
                            f++;
                            workSheet.Cells[f, c].Value = p.Key;
                            c++;

                        }

                        #region Cabecera lateral Totales

                        #endregion

                        f++;
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = "TOTAL AGRUPADAS:";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;



                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = "TOTAL AGORA:";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;
                        //workSheet.Cells[f, c].Value = total_agora;
                        //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        GuardaTotales(DateTime.Now, "TOTALES", "TOTAL AGORA", total_agora);


                        f++;
                        c = 1;

                        workSheet.Cells[f, c].Value = "TOTAL PASSPOOL 58:";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;
                        //workSheet.Cells[f, c].Value = total_pass_pool_58;
                        //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PASS POOL 58", total_pass_pool_58);
                        f++;
                        c = 1;

                        workSheet.Cells[f, c].Value = "TOTAL PASSPOOL 59:";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;
                        //workSheet.Cells[f, c].Value = total_pass_pool_59;
                        //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        GuardaTotales(DateTime.Now, "TOTALES", "TOTAL PASS POOL 59", total_pass_pool_59);
                        f++;
                        c = 1;

                        workSheet.Cells[f, c].Value = "TOTAL GESTIÓN PROPIA ATR:";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;
                        //workSheet.Cells[f, c].Value = total_gestion_propia_atr;
                        //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        GuardaTotales(DateTime.Now, "TOTALES", "TOTAL GESTION PROPIA ATR", total_gestion_propia_atr);



                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = "TOTAL MULTIPUNTO:";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        c++;
                        //workSheet.Cells[f, c].Value = total_multipuntos;
                        //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        GuardaTotales(DateTime.Now, "TOTALES", "TOTAL MULTIPUNTO", total_multipuntos);

                        c = 1;

                        firstOnly = true;
                        foreach (KeyValuePair<DateTime, List<EndesaEntity.facturacion.IRF_Totales_Hist>> p in dic_totales)
                        {
                            f = 1;

                            #region diferencia
                            if (firstOnly)
                            {
                                c++;
                                firstOnly = false;
                                workSheet.Cells[f, c].Value = "DIF (D-C)";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                f = 2;
                                workSheet.Cells[f, c].Value =
                                    GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "EMPRESA", workSheet.Cells[f, 1].Value.ToString()) -
                                    GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                                f = 3;
                                workSheet.Cells[f, c].Value =
                                    GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "EMPRESA", workSheet.Cells[f, 1].Value.ToString()) -
                                    GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                                f = 4;
                                workSheet.Cells[f, c].Value =
                                    GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "EMPRESA", workSheet.Cells[f, 1].Value.ToString()) -
                                    GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                                f++;
                                f++;
                                foreach (KeyValuePair<string, EndesaEntity.Agrupacion> pp in dic_agrup_tarifa)
                                {

                                    f++;
                                    workSheet.Cells[f, c].Value =
                                        GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TARIFA", pp.Key) -
                                        GetTotal(p.Key, "TARIFA", pp.Key);
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                }

                                f++;
                                f++;
                                foreach (KeyValuePair<string, EndesaEntity.Agrupacion> pp in dic_agrup_tipo_punto_producto)
                                {

                                    f++;
                                    workSheet.Cells[f, c].Value =
                                        GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TIPO PUNTO PRODUCTO", pp.Key) -
                                        GetTotal(p.Key, "TIPO PUNTO PRODUCTO", pp.Key);
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                }

                                f++;
                                f++;
                                foreach (KeyValuePair<string, EndesaEntity.Agrupacion> pp in dic_agrup_tipo_tarifa)
                                {

                                    f++;
                                    workSheet.Cells[f, c].Value =
                                        GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TIPO TARIFA", pp.Key) -
                                        GetTotal(p.Key, "TIPO TARIFA", pp.Key);
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                }

                                f++;
                                f++;
                                foreach (KeyValuePair<string, EndesaEntity.Agrupacion> pp in dic_agrup_exencion)
                                {

                                    f++;
                                    workSheet.Cells[f, c].Value =
                                        GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TIPO EXENCION", pp.Key) -
                                        GetTotal(p.Key, "TIPO EXENCION", pp.Key);
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                }


                                f++;
                                f++;
                                workSheet.Cells[f, c].Value =
                                    GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL AGRUPADAS") -
                                    GetTotal(p.Key, "TOTALES", "TOTAL AGRUPADAS");
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";


                                f++;
                                workSheet.Cells[f, c].Value =
                                    GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL AGORA") -
                                    GetTotal(p.Key, "TOTALES", "TOTAL AGORA");
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                                f++;
                                workSheet.Cells[f, c].Value =
                                    GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL PASS POOL 58") -
                                    GetTotal(p.Key, "TOTALES", "TOTAL PASS POOL 58");
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                                f++;
                                workSheet.Cells[f, c].Value =
                                    GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL PASS POOL 59") -
                                    GetTotal(p.Key, "TOTALES", "TOTAL PASS POOL 59");
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                                f++;
                                workSheet.Cells[f, c].Value =
                                    GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL GESTIÓN PROPIA ATR") -
                                    GetTotal(p.Key, "TOTALES", "TOTAL GESTIÓN PROPIA ATR");
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                                f++;
                                workSheet.Cells[f, c].Value =
                                    GetTotal(utilfecha.UltimoDiaHabilAnterior(p.Key), "TOTALES", "TOTAL MULTIPUNTO") -
                                    GetTotal(p.Key, "TOTALES", "TOTAL MULTIPUNTO");
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            }

                            #endregion

                            f = 1;

                            c++;
                            workSheet.Cells[f, c].Value = p.Key;
                            workSheet.Cells[f, c].Style.Font.Bold = true;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                            f++;
                            workSheet.Cells[f, c].Value = GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value = GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value = GetTotal(p.Key, "EMPRESA", workSheet.Cells[f, 1].Value.ToString());
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            f++;
                            foreach (KeyValuePair<string, EndesaEntity.Agrupacion> tarifa in dic_agrup_tarifa)
                            {

                                f++;
                                workSheet.Cells[f, c].Value = GetTotal(p.Key, "TARIFA", tarifa.Key);
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            }

                            f++;
                            f++;
                            foreach (KeyValuePair<string, EndesaEntity.Agrupacion> pp in dic_agrup_tipo_punto_producto)
                            {

                                f++;
                                workSheet.Cells[f, c].Value = GetTotal(p.Key, "TIPO PUNTO PRODUCTO", pp.Key);
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            }

                            f++;
                            f++;
                            foreach (KeyValuePair<string, EndesaEntity.Agrupacion> pp in dic_agrup_tipo_tarifa)
                            {

                                f++;
                                workSheet.Cells[f, c].Value = GetTotal(p.Key, "TIPO TARIFA", pp.Key);
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            }

                            f++;
                            f++;
                            foreach (KeyValuePair<string, EndesaEntity.Agrupacion> pp in dic_agrup_exencion)
                            {

                                f++;
                                workSheet.Cells[f, c].Value = GetTotal(p.Key, "TIPO EXENCION", pp.Key);
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            }

                            f++;
                            f++;
                            workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL AGRUPADAS");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";


                            f++;
                            workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL AGORA");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL PASS POOL 58");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL PASS POOL 59");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL GESTIÓN PROPIA ATR");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                            f++;
                            workSheet.Cells[f, c].Value = GetTotal(p.Key, "TOTALES", "TOTAL MULTIPUNTO");
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";

                        }



                        allCells = workSheet.Cells[1, 1, f, c];
                        allCells.AutoFitColumns();

                    }
                    excelPackage.SaveAs(nombreSalidaExcel);
                    excelPackage = null;

                    if (param.GetValue("EnviarMail") == "S")
                        EnvioCorreoInformeTipologias(nombreSalidaExcel.FullName);

                    ss_pp.Update_Fecha_Fin("Facturación", "Informe Inventario Tipologias PT", "Informe Inventario Tipologias PT");
                }
            }
            catch (Exception ex)
            {
                ficheroLog.addError("Inventario_por_Tipologias_v2: " + ex.Message);
            }


        }


        private int GetTotal(DateTime fecha_informe, string grupo, string concepto)
        {
            int total = 0;
            List<EndesaEntity.facturacion.IRF_Totales_Hist> o;
            if (dic_totales.TryGetValue(fecha_informe, out o))
                foreach (EndesaEntity.facturacion.IRF_Totales_Hist p in o)
                    if (p.grupo == grupo && p.concepto == concepto)
                        total = p.cantidad;

            return total;
        }

        private void GuardaTotales(DateTime fecha_informe, string grupo, string concepto, int cantidad)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;
            try
            {
                strSql = "replace into informe_inventario_tipologias_pt_totales_hist"
                    + " (fecha_informe, grupo, concepto, cantidad, last_update_date)"
                    + " values ('" + fecha_informe.ToString("yyyy-MM-dd") + "',"
                    + "'" + grupo + "',"
                    + "'" + concepto + "',"
                    + cantidad + ","
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);
                ficheroLog.AddError("InformeRevisionFacturas.GuardaTotales: " + e.Message);
            }
        }

        private Dictionary<DateTime, List<EndesaEntity.facturacion.IRF_Totales_Hist>> CargaTotales()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime fecha = new DateTime();

            Dictionary<DateTime, List<EndesaEntity.facturacion.IRF_Totales_Hist>> d =
                new Dictionary<DateTime, List<EndesaEntity.facturacion.IRF_Totales_Hist>>();

            try
            {

                strSql = "select fecha_informe, grupo, concepto, cantidad"
                    + " from informe_inventario_tipologias_pt_totales_hist"
                    + " where fecha_informe > '" + DateTime.Now.AddDays(-31).ToString("yyyy-MM-dd") + "'"
                    + " order by fecha_informe desc";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.IRF_Totales_Hist c = new EndesaEntity.facturacion.IRF_Totales_Hist();
                    fecha = Convert.ToDateTime(r["fecha_informe"]);
                    c.grupo = r["grupo"].ToString();
                    c.concepto = r["concepto"].ToString();
                    c.cantidad = Convert.ToInt32(r["cantidad"]);

                    List<EndesaEntity.facturacion.IRF_Totales_Hist> o;
                    if (!d.TryGetValue(fecha, out o))
                    {
                        o = new List<EndesaEntity.facturacion.IRF_Totales_Hist>();
                        o.Add(c);
                        d.Add(fecha, o);
                    }
                    else
                        o.Add(c);

                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        private void EnvioCorreoInformeTipologias(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();
            

            try
            {
                Console.WriteLine("Enviando correo Informe de Tipologias");

                string from = param.GetValue("Buzon_envio_email");
                string to = param.GetValue("email_tipologias_para");
                string cc = param.GetValue("email_tipologias_copia");
                string subject = param.GetValue("email_asunto_tipologias") + " " + DateTime.Now.ToString("dd/MM/yyyy");
                

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   Adjuntamos el informe de contratos por tipologías.");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("P.D:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   - La fecha de actualización de la extracción contratos BTN es: "
                    + ss_pp.GetFecha_FinProceso("Contratación", "contratos_PS_Portugal_BTN", "2_Importar_Contratos_BTN").ToString("dd/MM/yyyy")).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   - La fecha de actualización de la extracción complementos BTN es: "
                   + ss_pp.GetFecha_FinProceso("Contratación", "contratos_complementos_PS_Portugal_BTN", "2_Importar_Contratos_Complementos_BTN").ToString("dd/MM/yyyy")).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   - La fecha de actualización de la extracción contratos BTE es: "
                    + ss_pp.GetFecha_FinProceso("Contratación", "contratos_PS_Portugal_BTE", "2_Importar_Contratos_BTE").ToString("dd/MM/yyyy")).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   - La fecha de actualización de la extracción complementos BTE es: "
                   + ss_pp.GetFecha_FinProceso("Contratación", "contratos_complementos_PS_Portugal_BTE", "2_Importar_Contratos_Complementos_BTE").ToString("dd/MM/yyyy")).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   - La fecha de actualización de la extracción contratos MT es: "
                    + ss_pp.GetFecha_FinProceso("Contratación", "contratos_PS_Portugal_MT", "2_Importar_Contratos_MT").ToString("dd/MM/yyyy")).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   - La fecha de actualización de la extracción complementos MT es: "
                   + ss_pp.GetFecha_FinProceso("Contratación", "contratos_complementos_PS_Portugal_MT", "2_Importar_Contratos_Complementos_MT").ToString("dd/MM/yyyy")).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);


                //EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();                
                //mes.SendMailNoHTML(from, to, cc, subject, textBody.ToString(), archivo);

                EndesaBusiness.office.SendMail mes = new office.SendMail(from);
                List<string> lista_para = to.Split(';').ToList();
                List<string> lista_cc = cc.Split(';').ToList();
                List<string> lista_adjuntos = archivo.Split(';').ToList();

                mes.para = lista_para;
                mes.cc = lista_cc;
                mes.asunto = subject;
                mes.htmlCuerpo = textBody.ToString();
                mes.adjuntos = lista_adjuntos;

                if (param.GetValue("EnviarMail") == "S")
                    mes.Send();
                else
                    mes.Save();


                ficheroLog.Add("Correo enviado desde: " + param.GetValue("Buzon_envio_email"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreoInformeTipologias: " + e.Message);
            }
        }

        public void Importacion_Ajustes_Exenciones()
        {
            EndesaBusiness.utilidades.Fechas ff = new utilidades.Fechas();
            StringBuilder sb = new StringBuilder();
            string md5 = "";

            DateTime ultima_ejecucion = new DateTime();


            try
            {
                Console.WriteLine("Última Ejecución: "
                    + ss_pp.GetFecha_FinProceso("Facturación", "Informe Tipologías PT", "2_Importación").ToString("dd/MM/yyyy"));

                ultima_ejecucion =
                ss_pp.GetFecha_FinProceso("Facturación", "Informe Tipologías PT", "2_Importación");

                Console.WriteLine(ultima_ejecucion.ToString("dd/MM/yyyy") + " > " + DateTime.Now.Date.ToString("dd/MM/yyyy"));

                if (ultima_ejecucion < DateTime.Now.Date)
                {


                    string archivo = param.GetValue("Ubicacion_Informes")   
                        + param.GetValue("prefijo_archivo_exenciones").Replace("*", ff.UltimoDiaHabil().ToString("yyMMdd"))
                        + ".txt";

                    DescargaExtractorAjustes(ff.UltimoDiaHabil());

                    FileInfo fileInfo = new FileInfo(archivo);
                    if (fileInfo.Exists)
                    {

                        md5 = utilidades.Fichero.checkMD5(archivo).ToString();
                        if (fileInfo.Length > 0 && (md5 != param.GetValue("MD5_archivo_exenciones")))
                        {


                            ImportarArchivoExenciones(archivo);

                            param.code = "MD5_archivo_exenciones";
                            param.from_date = new DateTime(2022, 01, 01);
                            param.to_date = new DateTime(4999, 12, 31);
                            param.value = md5;
                            param.Save();
                        }

                    }
                    
                }
            
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                ficheroLog.AddError(ex.Message);
            }
        }

        private void DescargaExtractorAjustes(DateTime fecha)
        {
            try
            {

                ficheroLog.Add("Ejecutando extractor: " + param.GetValue("Extractor_Ajustes_Exenciones", DateTime.Now, DateTime.Now));

                ss_pp.Update_Fecha_Inicio("Facturación", "Informe Tipologías PT", "1_Ejecución Extractor");

                utilidades.Fichero.EjecutaComando(param.GetValue("Extractor_Ajustes_Exenciones"), fecha.ToString("yyMMdd"));

                ss_pp.Update_Fecha_Fin("Facturación", "Informe Tipologías PT", "1_Ejecución Extractor");

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Descarga --> " + e.Message);
            }


        }

        private void ImportarArchivoExenciones(string archivo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            string line = "";
            string[] campos;
            int i = 0;
            int c = 0;
            string strSql = "";
            int total_registros = 0;


            try
            {

                ss_pp.Update_Fecha_Inicio("Facturación", "Informe Tipologías PT", "2_Importación");

                strSql = "delete from informe_inventario_tipologias_pt_exenciones";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Importando EXENCIONES");
                Console.WriteLine("=====================");

                ficheroLog.Add("Importando archivo " + archivo);

                System.IO.StreamReader file = new System.IO.StreamReader(archivo, System.Text.Encoding.GetEncoding(1252));
                while ((line = file.ReadLine()) != null)
                {
                    campos = line.Split('$');
                    i++;
                    c = 0;

                    if (i > 1)
                    {
                        total_registros++;

                        if (firstOnly)
                        {
                            sb.Append("replace into informe_inventario_tipologias_pt_exenciones");
                            sb.Append(" (tipo_ajuste, tipo_idenficador_fiscal, identificador_fiscal,");
                            sb.Append(" cups13, cups22, razon_social, nombre, fecha_inicio_vigencia,");
                            sb.Append("fecha_fin_vigencia, num_dias_hasta_caducar, porcentaje_exento,");
                            sb.Append("porcentaje_aplicacion, codigo_tarjeta, articulo_cie, created_by, created_date");
                            sb.Append(" ) values ");
                            firstOnly = false;
                        }

                        sb.Append("('").Append(FuncionesTexto.RT(campos[c])).Append("',"); c++; // tipo_ajuste
                        sb.Append("'").Append(FuncionesTexto.RT(campos[c])).Append("',"); c++; // tipo_idenficador_fiscal                  
                        sb.Append("'").Append(FuncionesTexto.RT(campos[c])).Append("',"); c++; // identificador_fiscal
                                                                                   // 
                        sb.Append("'").Append(FuncionesTexto.RT(campos[c]).Trim()).Append("',"); c++; // cups13
                        sb.Append("'").Append(FuncionesTexto.RT(campos[c]).Trim()).Append("',"); c++; // cups22
                        sb.Append("'").Append(FuncionesTexto.RT(campos[c])).Append("',"); c++; // razon_social
                        sb.Append("'").Append(FuncionesTexto.RT(campos[c])).Append("',"); c++; // nombre
                        sb.Append(CF(campos[c])).Append(","); c++; // fecha_inicio_vigencia

                        sb.Append(CF(campos[c])).Append(","); c++; // fecha_fin_vigencia
                        sb.Append("'").Append(FuncionesTexto.RT(campos[c])).Append("',"); c++; // num_dias_hasta_caducar
                        sb.Append(CN(campos[c])).Append(","); c++; // porcentaje_exento

                        sb.Append(CN(campos[c])).Append(","); c++; // porcentaje_aplicacion
                        sb.Append("'").Append(FuncionesTexto.RT(campos[c])).Append("',"); c++; // codigo_tarjeta     
                        sb.Append("'").Append(FuncionesTexto.RT(campos[c])).Append("',"); c++; // articulo_cie                    
                        sb.Append("'").Append(System.Environment.UserName).Append("',"); // created_by
                        sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'"); // created_date
                        sb.Append("),");

                        if (i == 250)
                        {

                            Console.CursorLeft = 0;
                            Console.Write(total_registros.ToString("N0"));

                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.CON);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            i = 0;
                        }
                    }

                }
                file.Close();

                if (i > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                }

                ficheroLog.Add("Fin importación");
                ficheroLog.Add("Se han importado " + total_registros.ToString("N0") + " registros");
                FileInfo archivo_info = new FileInfo(archivo);
                archivo_info.Delete();

                ss_pp.Update_Fecha_Fin("Facturación", "Informe Tipologías PT", "2_Importación");
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("ImportarArchivoContratos: " + ex.Message);
                Console.WriteLine("ImportarArchivoContratos: " + ex.Message);
            }
        }

        private Dictionary<string, EndesaEntity.contratacion.Exenciones> CargaExenciones()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            Dictionary<string, EndesaEntity.contratacion.Exenciones> d =
                new Dictionary<string, EndesaEntity.contratacion.Exenciones>();

            try
            {

                strSql = "SELECT e.cups22, e.tipo_ajuste FROM informe_inventario_tipologias_pt_exenciones e"
                    + " WHERE e.fecha_inicio_vigencia <= NOW() AND e.fecha_fin_vigencia >= NOW()";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.Exenciones c =
                        new EndesaEntity.contratacion.Exenciones();

                    if (r["cups22"] != System.DBNull.Value)
                        c.cups22 = r["cups22"].ToString();

                    if (r["tipo_ajuste"] != System.DBNull.Value)
                        c.tipo_ajuste = r["tipo_ajuste"].ToString();

                    EndesaEntity.contratacion.Exenciones o;
                    if (!d.TryGetValue(c.cups22, out o))
                        d.Add(c.cups22, c);

                }
                db.CloseConnection();

                return d;
            }
            catch(Exception ex)
            {
                return null;
            }
        }


        private string Get_TipoAjuste(string cups22)
        {
            EndesaEntity.contratacion.Exenciones o;
            if (dic_exenciones.TryGetValue(cups22, out o))
                return o.tipo_ajuste;
            else
                return "";
        }


        private string CF(string t)
        {
            if (t.Trim() == "00000000" || t.Trim() == "")
                return "null";
            else
                return "'" + t.Substring(0, 4)
                    + "-" + t.Substring(4, 2)
                    + "-" + t.Substring(6, 2) + "'";
        }

        private string CN(string t)
        {
            if (t.Trim() == "00000000" || t.Trim() == "")
                return "'null'";
            else
                return "'" + t.Trim() + "'";
        }

        public string CDouble(String t)
        {
            t = t.Trim();
            if (t == "")
            {
                return "null";
            }
            else
            {
                t = t.Replace(" ", "");
                t = t.Replace("+", string.Empty);
                t = t.Replace("----------", string.Empty);
                t = t.Replace(".", string.Empty);
                t = t.Replace(",", ".");

                if (t == "")
                {
                    t = "null";
                }
            }

            return t;
        }

    }
}

