﻿using EndesaBusiness.contratacion.eexxi;
using EndesaBusiness.factoring;
using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using EndesaEntity.cnmc.V21_2019_12_17;
using EndesaEntity.cnmc.V30_2022_21_01;
using EndesaEntity.contratacion.gas;
using EndesaEntity.extrasistemas;
using Microsoft.Graph;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using Org.BouncyCastle.Ocsp;
using System;
using System.Globalization;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using Telegram.Bot.Types;
using System.Diagnostics.Eventing.Reader;
using Microsoft.Win32;

namespace EndesaBusiness.xml 
{
    public class Extrasistemas
    {
        public utilidades.Param param;
        //irh  public
        public utilidades.Param param_cnmc;
        List<EndesaEntity.extrasistemas.Global> lista_datos_excel_extrasistemas;
        Dictionary<string,EndesaEntity.xml.Plantilla_Extrasistemas_total_registros> dic_total_registros;
        public List<string> lista_log { get; set; }
        cnmc.CNMC cnmc;

        public Extrasistemas()
        {
            param = new utilidades.Param("extrasistemas_param", servidores.MySQLDB.Esquemas.CON);
            param_cnmc = new utilidades.Param("cnmc_param", servidores.MySQLDB.Esquemas.CON);
            lista_datos_excel_extrasistemas = new List<EndesaEntity.extrasistemas.Global>();
            dic_total_registros = Inicializa_Lista_Totales();
            lista_log = new List<string>();
            cnmc = new cnmc.CNMC();
        }

        public void CargaExcel(string fichero)
        {
            ProcesaExcel(fichero);
            //RellenaInventario(dic);
            //Proceso(mensual);
            //GeneraExcelResultados(fichero);
        }

        private void ProcesaExcel(string fichero)
        {
            int f = 1;
            int c = 1;
            int total_hojas_excel = 0;
            bool firstOnly = true;
            string cabecera = "";
            FileStream fs;
            ExcelPackage excelPackage;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
          //  EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 c101 =
          //      new EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101();

            string tipoDocAportado = "";
            string direccionUrl = "";

            bool registro_valido = false;

            int ad_total = 0;
            int acc_total = 0;
            int acc_cambios_total = 0;
            int mod_total = 0;
            int baja_total = 0;

            int ad_ok = 0;
            int acc_ok = 0;
            int acc_cambios_ok = 0;
            int mod_ok = 0;
            int baja_ok = 0;

            string fecha = "";
            EndesaEntity.xml.Plantilla_Extrasistemas_total_registros o;

            try
            {
                lista_log.Clear();

                FileInfo file = new FileInfo(fichero);
                fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                excelPackage = new ExcelPackage(fs);
                total_hojas_excel = excelPackage.Workbook.Worksheets.Count();
                var workSheet = excelPackage.Workbook.Worksheets.First();
                for (int hoja = 0; hoja < total_hojas_excel; hoja++)
                {
                    workSheet = excelPackage.Workbook.Worksheets[hoja];
                    firstOnly = true;


                    if (workSheet.Name == "AD")
                    // A301 Alta Directa  --  tabla proceso A3 - Altas.xsd
                    {
                        //NO PROCESAMOS ESTA PESTAÑA HASTA FINALIZAR IMPLEMENTACION
                        //FORZAMOS BREAK
                        //lista_log.Add("La pestaña AD  no se procesará, está deshabilitada temporalmente");
                        //continue;
                        //

                        f = 2; // Porque la primera fila es la cabecera
                        for (int i = 1; i < 1000000; i++)
                        {
                            registro_valido = true;
                            f++;
                            if (workSheet.Cells[f, 3].Value == null)
                                break;

                            ad_total++;

                            EndesaEntity.extrasistemas.Global g = new EndesaEntity.extrasistemas.Global();
                            g.fichero = file.Name;
                            g.hoja = workSheet.Name;



                            //irh 28-01-25   _   EndesaBusiness / xml /  Extrasistemas.cs
                            //  CodigoREEEmpresaEmisora - Obligatorio X(4) - STRING NUMERICO  - 

                            if (workSheet.Cells[f, 1].Value != null && Convert.ToString(workSheet.Cells[f, 1].Value).Length >= 4)
                            {
                                if (int.TryParse(Convert.ToString(workSheet.Cells[f, 1].Value).Substring(0, 4), out _))

                                    g.empresa_emisora = Convert.ToString(workSheet.Cells[f, 1].Value).Substring(0, 4);
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " Empresa_emisora: el campo código empresa emisora no tiene un formato válido.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " Empresa_emisora: el campo código empresa emisora es nulo o no tiene una longitud válida X(4).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //  old :   g.empresa_emisora = Convert.ToString(workSheet.Cells[f, 1].Value).Substring(0, 4);


                            //  Distribuidora
                            int celdaDistribuidora = 2;  // Variable para la celda de Distribuidora

                            // Validar Distribuidora (4 caracteres)
                            if (workSheet.Cells[f, celdaDistribuidora].Value != null && Convert.ToString(workSheet.Cells[f, celdaDistribuidora].Value).Length >= 4)
                            {
                                if (int.TryParse(Convert.ToString(workSheet.Cells[f, celdaDistribuidora].Value).Substring(0, 4), out _))
                                {
                                    g.distribuidora = Convert.ToString(workSheet.Cells[f, celdaDistribuidora].Value).Substring(0, 4);
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Distribuidora: el campo código empresa destino (distribuidora) no tiene un formato válido.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Distribuidora: el campo código empresa destino (distribuidora) es nulo o no tiene una longitud válida X(4).");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            // old :  g.distribuidora = Convert.ToString(workSheet.Cells[f, 2].Value);

                            //  
                            //  cups22
                            int celdaCUPS = 3;   // Variable para la celda de CUPS
                                                 // Validar CUPS (22 caracteres)
                            if (workSheet.Cells[f, celdaCUPS].Value != null && Convert.ToString(workSheet.Cells[f, celdaCUPS].Value).Length >= 22)
                            {
                                g.cups22 = Convert.ToString(workSheet.Cells[f, celdaCUPS].Value).Substring(0, 22);
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> CUPS: el campo CUPS es nulo o no tiene una longitud válida X(22).");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            // old : g.cups22 = Convert.ToString(workSheet.Cells[f, 3].Value);

                            #region DatosSolicitudAD_A3
                            //     cnae 
                            int celdaCNAE = 4;   // Variable para la celda de CNAE
                            if (workSheet.Cells[f, celdaCNAE].Value != null && Convert.ToString(workSheet.Cells[f, celdaCNAE].Value).Length >= 4)
                            {
                                g.cnae = Convert.ToString(workSheet.Cells[f, celdaCNAE].Value).Substring(0, 4);
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> CNAE: el campo CNAE es nulo o no tiene una longitud válida X(4).");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //  old  g.cnae = Convert.ToString(workSheet.Cells[f, 4].Value);

                            // IndActivacion fila  5 : cotiene   F fecha fija ,   A cuanto antes o  L - proxima lectura
                            //  ojo accedo a un diccionario_indicativo activacion 
                            int celdaIndActivacion = 5;  // Variable para la celda de IndActivacion

                            if (workSheet.Cells[f, celdaIndActivacion].Value != null && Convert.ToString(workSheet.Cells[f, celdaIndActivacion].Value).Length >= 1)
                            {
                                string valorIndActivacion = Convert.ToString(workSheet.Cells[f, celdaIndActivacion].Value).Substring(0, 1).ToUpper();

                                if (cnmc.dic_indicativo_activacion.ContainsValue(valorIndActivacion))
                                {
                                    g.ind_activacion = valorIndActivacion;
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> IndActivacion: el campo IndActivacion no contiene un valor válido.");
                                    // lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> IndActivacion: el campo IndActivacion es nulo o no tiene una longitud válida X(1).");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            // old    g.ind_activacion = Convert.ToString(workSheet.Cells[f, 5].Value).Substring(0, 1);


                            //  fecha_activacion - Condicionada a valor IndActivacion == F - Fecha(AAAA-MM-DD)
                            int celdaFechaActivacion = 6;  // Variable para la celda de fecha_activacion
                            // fecha_activacion - Condicionada a que IndActivacion sea "F" - Fecha en formato AAAA-MM-DD
                            if (g.ind_activacion == "F")
                            {
                                var valorFecha = workSheet.Cells[f, celdaFechaActivacion].Value;

                                if (valorFecha != null)
                                {
                                    string fechaStr = Convert.ToString(valorFecha).Trim();

                                    if (DateTime.TryParse(fechaStr, out DateTime fechaParsed))
                                    {
                                        // Verifica que el formato sea exactamente AAAA-MM-DD
                                        if (fechaParsed.ToString("yyyy-MM-dd") == fechaStr)
                                        {
                                            g.fecha_activacion = fechaParsed;
                                        }
                                        else
                                        {
                                            lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> FechaPrevistaAccion: el campo fecha_activación no tiene el formato AAAA-MM-DD.");
                                            lista_log.Add(Environment.NewLine);
                                            registro_valido = false;
                                        }
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> FechaPrevistaAccion: el campo fecha_activación no es una fecha válida.");
                                        lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> FechaPrevistaAccion: el campo fecha_activación está vacío.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }


                            #endregion
                            #region ContratoAD_A3
                            //  tipo de contrato atr  , fila 7  -TipoContratoATR  - S- x 2 -  tabla 9  --   ver tabla diccionario ???-----------------------------------
                            // dic_contrato_atr = Carga_Tabla_CNMC("cnmc_p_tipo_contrato_atr");

                            if (workSheet.Cells[f, 7].Value != null && Convert.ToString(workSheet.Cells[f, 7].Value).Length >= 2)
                            {
                                //  ojo la clave en el diccionario esta al revez en la  descripcion  .----
                                if (cnmc.dic_contrato_atr.ContainsValue(Convert.ToString(workSheet.Cells[f, 7].Value).Substring(0, 2).ToUpper()))
                                {
                                    g.tipo_contrato_atr = Convert.ToString(workSheet.Cells[f, 7].Value).Substring(0, 2).ToUpper();

                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " tipo_contrato_at: el campo tipo_contrato_at no contiene un valor válido.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + " tipo_contrato_at : el campo tipo_contrato_at es nulo o no tiene una longitud válida X2).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //  old   g.tipo_contrato_atr = Convert.ToString(workSheet.Cells[f, 7].Value);

                            // fecha Finalizacion , fial 8     aaaa-mm-dd
                            // Obtener y limpiar el valor de la celda 8
                            string fecha8 = workSheet.Cells[f, 8]?.Value?.ToString()?.Trim();

                            if (!string.IsNullOrEmpty(fecha8) && DateTime.TryParseExact(fecha8, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fechaFinalizacion))
                            {
                                g.fecha_finalizacion = fechaFinalizacion;

                                // Validar si la fecha de finalización debe estar informada
                                if (new[] { "02", "03", "09" }.Contains(g.tipo_contrato_atr) && g.fecha_finalizacion == null)
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> Fila: {f} ➝ La Fecha Finalización es obligatoria para el tipo de contrato '{g.tipo_contrato_atr}'. Rechazo '08'.");
                                    registro_valido = false;
                                }

                                // Validar que la fecha no sea superior a 1 año después de la fecha de activación
                                //if (g.fecha_activacion != null && fechaFinalizacion > g.fecha_activacion.Value.AddYears(1))
                                //{
                                //    lista_log.Add($"Hoja: {g.hoja} --> Fila: {f} ➝ La Fecha Finalización '{fechaFinalizacion:yyyy-MM-dd}' excede un año desde la Fecha de Activación '{g.fecha_activacion:yyyy-MM-dd}'. Rechazo '08'.");
                                //     registro_valido = false;
                                // }

                                // Validar que la fecha de finalización no sea menor a la fecha de activación
                                if (g.fecha_activacion != null && fechaFinalizacion < g.fecha_activacion)
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> Fila: {f} ➝ La Fecha Finalización '{fechaFinalizacion:yyyy-MM-dd}' es menor a la Fecha de Activación '{g.fecha_activacion:yyyy-MM-dd}'. Rechazo '08'.");
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($" Hoja: {g.hoja} --> Fila: {f} ➝ Fecha Finalización: '{fecha8 ?? "NULO"}' no es válida (debe ser AAAA-MM-DD).");
                                //lista_log.Add(System.Environment.NewLine);  // Agrega una línea en blanco al log
                                registro_valido = false;
                            }


                            #region autoconsumocau
                            // Tipo de autoconumo    fila 9 - 
                            if (workSheet.Cells[f, 9].Value != null && Convert.ToString(workSheet.Cells[f, 9].Value).Length >= 2)
                            {
                                //  ojo la clave es la descripcion  .---- tipo autoconsumo   ojo contaisvalue --- y no cnmc.dic_tarifa_atr.ContainsKey  -  (containsKey)
                                if (cnmc.dic_autoconsumo.ContainsValue(Convert.ToString(workSheet.Cells[f, 9].Value).Substring(0, 2).ToUpper()))
                                {
                                    // if (Convert.ToString(workSheet.Cells[f, 9].Value).Substring(0, 2).ToUpper() == "00")
                                    //     g.tipo_autoconsumo = null;
                                    //  else

                                    g.tipo_autoconsumo = Convert.ToString(workSheet.Cells[f, 9].Value).Substring(0, 2).ToUpper();

                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " tipo_autoconsumo: el campo tipo_autoconsumo no contiene un valor válido.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + " tipo_autoconsumo : el campo tipo_autoconsumo es nulo o no tiene una longitud válida X2).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;

                            }
                            // old  g.tipo_autoconsumo = Convert.ToString(workSheet.Cells[f, 9].Value);
                            //
                            // Verificar si algún CAU tiene TipoAutoconsumo diferente de "00" y "0C"
                            bool informarCampos = g.tipo_autoconsumo != "00" && g.tipo_autoconsumo != "0C";

                            if (informarCampos)
                            {

                                // Crear variable para almacenar el valor de la celda 10
                                int celdatipocuops = 10;
                                var valorCeldatc = workSheet.Cells[f, celdatipocuops].Value;

                                if (valorCeldatc != null && Convert.ToString(valorCeldatc).Length >= 2)
                                {
                                    // Se obtiene el valor de las primeras dos letras en mayúsculas
                                    string celdatipocups = Convert.ToString(valorCeldatc).Substring(0, 2).ToUpper();

                                    // Asignar el valor directamente
                                    g.tipocups = celdatipocups;
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} tipocups: el campo tipocups es nulo o no tiene una longitud válida.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                                //

                                int celdaCau = 11;
                                var valorCeldaCau = workSheet.Cells[f, celdaCau].Value;

                                if (valorCeldaCau != null && Convert.ToString(valorCeldaCau).Length == 26)
                                {
                                    // Asignar el valor directamente
                                    g.cau = Convert.ToString(valorCeldaCau);
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} cau: el campo cau es nulo o no tiene una longitud válida.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                                //
                                int celdaTiposubseccion = 12;
                                var valorCeldaTiposubseccion = workSheet.Cells[f, celdaTiposubseccion].Value;

                                if (valorCeldaTiposubseccion != null && Convert.ToString(valorCeldaTiposubseccion).Length >= 2)
                                {
                                    // Se obtiene el valor de las primeras dos letras en mayúsculas
                                    string tiposubseccionValor = Convert.ToString(valorCeldaTiposubseccion).Substring(0, 2).ToUpper();

                                    // Asignar el valor directamente
                                    g.tipo_subseccion = tiposubseccionValor;
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} tiposubseccion: el campo tiposubseccion es nulo o no tiene una longitud válida.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                                //
                                int celdaColectivo = 13;
                                var valorCeldaColectivo = workSheet.Cells[f, celdaColectivo].Value;

                                if (valorCeldaColectivo != null && (Convert.ToString(valorCeldaColectivo) == "S" || Convert.ToString(valorCeldaColectivo) == "N"))
                                {
                                    g.colectivo = Convert.ToString(valorCeldaColectivo);
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} colectivo: el campo colectivo es nulo o no tiene un valor válido.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                                //
                                int celdaPotInstaladaGen = 14;

                                var valorCeldaPotInstaladaGen = workSheet.Cells[f, celdaPotInstaladaGen].Value;

                                if (valorCeldaPotInstaladaGen != null && long.TryParse(Convert.ToString(valorCeldaPotInstaladaGen), out long potInstaladaGenValor) && Convert.ToString(valorCeldaPotInstaladaGen).Length <= 14)
                                {
                                    // Asignar el valor directamente
                                    g.potinstaladagen = (int)potInstaladaGenValor;
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} potInstaladaGen: el campo potInstaladaGen es nulo, no es numérico o excede las 14 posiciones.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                                //
                                int celdaTipoInstalacion = 15;
                                var valorCeldaTipoInstalacion = workSheet.Cells[f, celdaTipoInstalacion].Value;

                                if (valorCeldaTipoInstalacion != null && Convert.ToString(valorCeldaTipoInstalacion).Length >= 2)
                                {
                                    string tipoInstalacionValor = Convert.ToString(valorCeldaTipoInstalacion).Substring(0, 2).ToUpper();

                                    if (g.tipo_autoconsumo == "11")
                                    {
                                        // Si TipoAutoconsumo es "11", solo se permite "01" o "02"
                                        if (tipoInstalacionValor == "01" || tipoInstalacionValor == "02")
                                        {
                                            g.tipoinstalacion = tipoInstalacionValor;
                                        }
                                        else
                                        {
                                            lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} TipoInstalacion: cuando TipoAutoconsumo es '11', solo puede ser '01' o '02'. Se recibió '{tipoInstalacionValor}'.");
                                            //lista_log.Add(System.Environment.NewLine);
                                            registro_valido = false;
                                        }
                                    }
                                    else
                                    {
                                        // Si TipoAutoconsumo NO es "11", cualquier valor informado en TipoInstalacion es válido
                                        g.tipoinstalacion = tipoInstalacionValor;
                                    }
                                }
                                else
                                {
                                    // Si TipoAutoconsumo es "11", TipoInstalacion es obligatorio
                                    if (g.tipo_autoconsumo == "11")
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} TipoInstalacion: cuando TipoAutoconsumo es '11', el campo TipoInstalacion es obligatorio ('01' o '02').");
                                        //lista_log.Add(System.Environment.NewLine);
                                        registro_valido = false;
                                    }
                                    else
                                    {
                                        // Si el TipoAutoconsumo NO es "11", TipoInstalacion puede estar vacío (nulo)
                                        g.tipoinstalacion = null;
                                    }
                                }
                                //
                                int celdaSsaa = 16;
                                var valorCeldaSsaa = workSheet.Cells[f, celdaSsaa].Value;

                                if (valorCeldaSsaa != null && (Convert.ToString(valorCeldaSsaa) == "S" || Convert.ToString(valorCeldaSsaa) == "N"))
                                {
                                    // Asignar el valor directamente
                                    g.ssaa = Convert.ToString(valorCeldaSsaa);
                                }
                                else if (g.tipocups == "01")
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} ssaa: el campo ssaa es obligatorio y debe contener 'S' o 'N' cuando TipoCups es '01'.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} ssaa: el campo ssaa es nulo o no tiene un valor válido.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                                //
                                int celdaUnicoContrato = 17;
                                var valorCeldaUnicoContrato = workSheet.Cells[f, celdaUnicoContrato].Value;
                                if (g.ssaa == "S")
                                {
                                    if (valorCeldaUnicoContrato != null && (Convert.ToString(valorCeldaUnicoContrato) == "S" || Convert.ToString(valorCeldaUnicoContrato) == "N"))
                                    {

                                        g.unicocontrato = Convert.ToString(valorCeldaUnicoContrato);
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} unicocontrato: el campo unicocontrato es obligatorio y debe contener 'S' o 'N' cuando SSAA es 'S'.");
                                        //lista_log.Add(System.Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    // Asignar el valor directamente si SSAA no es 'S'
                                    g.unicocontrato = Convert.ToString(valorCeldaUnicoContrato);
                                }
                            }
                            #endregion
                            //
                            #endregion
                            //
                            #region potencia
                            //     Tarifa    
                            if (workSheet.Cells[f, 18].Value != null && Convert.ToString(workSheet.Cells[f, 18].Value).Length >= 3)
                            {
                                //  ojo la clave es la descripcion  .---- tarifa  ojo contaisvalue ---no   (containsKey)
                                if (cnmc.dic_tarifa_atr.ContainsValue(Convert.ToString(workSheet.Cells[f, 18].Value).Substring(0, 3).ToUpper()))
                                {
                                    g.tarifa = Convert.ToString(workSheet.Cells[f, 18].Value).Substring(0, 3).ToUpper();

                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " tarifa atr: el campo tarifa atr no contiene un valor válido.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + " tarifa atr : el campo tarifa atr  es nulo o no tiene una longitud válida X3).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;

                            }

                            //   old     g.tarifa = Convert.ToString(workSheet.Cells[f, 10].Value).Substring(0,3);

                            //   potencia  en Watios  maximo   6     fila 11
                            // Validación de las potencias en vatios (máximo 6 valores)
                            int potenciasInformadas = 0;
                            bool esTarifa20TD = (g.tarifa == "2.0TD");

                            for (int pot = 0; pot < 6; pot++)
                            {
                                var potenciaValue = workSheet.Cells[f, 19 + pot].Value;
                                if (potenciaValue != null && int.TryParse(Convert.ToString(potenciaValue), out int potencia))
                                {
                                    if (potencia >= 0)
                                    {
                                        g.potencias[pot] = potencia;
                                        potenciasInformadas++;
                                    }
                                    else
                                    {
                                        lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Potencia " + (pot + 1) + ": el valor debe ser un número positivo.");
                                        //lista_log.Add(System.Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Potencia " + (pot + 1) + ": el valor no es un número válido o está en blanco.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }

                            // Verificación de que al menos 3 potencias estén informadas
                            if (potenciasInformadas < 3)
                            {
                                lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Deben estar informadas al menos 3 potencias.");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }

                            // Verificación del orden creciente de las potencias (excepto para tarifa 2.0TD)
                            if (!esTarifa20TD)
                            {
                                for (int pot = 1; pot < 6; pot++)
                                {
                                    if (g.potencias[pot] < g.potencias[pot - 1])
                                    {
                                        lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Las potencias deben ser crecientes. Potencia " + (pot + 1) + " (" + g.potencias[pot] + ") es menor que Potencia " + pot + " (" + g.potencias[pot - 1] + ").");
                                        //lista_log.Add(System.Environment.NewLine);
                                        registro_valido = false;
                                        break;
                                    }
                                }
                            }
                            // old
                            //for (int pot = 0; pot < 6; pot++)
                            //{
                            //    if(Convert.ToString(workSheet.Cells[f, 11 + pot].Value) != "")
                            //        g.potencias[pot] = Convert.ToInt32(workSheet.Cells[f, 11 + pot].Value);
                            //}
                            //
                            //
                            #endregion

                            #region ClienteAD_A3
                            if (workSheet.Cells[f, 25].Value != null)
                            {
                                // Normalizar el valor de la celda
                                string descripcion = Convert.ToString(workSheet.Cells[f, 25].Value)?.Trim().ToUpperInvariant();

                                if (!string.IsNullOrEmpty(descripcion))
                                {
                                    // Normalizar claves del diccionario eliminando espacios y asegurando mayúsculas
                                    var dic_normalizado = cnmc.dic_identificador
                                        .ToDictionary(kv => kv.Key.Trim().ToUpperInvariant(), kv => kv.Value);

                                    // Buscar la clave en el diccionario
                                    if (!dic_normalizado.TryGetValue(descripcion, out string valorEncontrado))
                                    {
                                        // Intento adicional: Buscar con coincidencias parciales si las claves son listas
                                        foreach (var kvp in dic_normalizado)
                                        {
                                            var clavesDiccionario = kvp.Key.Split(new[] { ',', '-', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                                           .Select(s => s.Trim().ToUpperInvariant());

                                            var clavesEntrada = descripcion.Split(new[] { ',', '-', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                                           .Select(s => s.Trim().ToUpperInvariant());

                                            if (clavesDiccionario.Intersect(clavesEntrada).Any()) // Coincidencia parcial
                                            {
                                                valorEncontrado = kvp.Value;
                                                break;
                                            }
                                        }
                                    }

                                    // Verificar si encontramos el valor
                                    if (!string.IsNullOrEmpty(valorEncontrado))
                                    {
                                        if (valorEncontrado.Length >= 2)
                                        {
                                            g.tipo_identificador = valorEncontrado.Substring(valorEncontrado.Length - 2).ToUpper();
                                        }
                                        else
                                        {
                                            lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El valor asociado a '{descripcion}' en el diccionario es inválido ({valorEncontrado}).");
                                            lista_log.Add(Environment.NewLine);
                                            registro_valido = false;
                                        }
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: La descripción '{descripcion}' no se encontró en el diccionario.");
                                        lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El valor en la celda es inválido (vacío o solo espacios).");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El campo es nulo.");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }

                            // old   g.tipo_identificador = Convert.ToString(workSheet.Cells[f, 17].Value);

                            //  numero de identificador  
                            g.n_identificador = Convert.ToString(workSheet.Cells[f, 26].Value);

                            //  tipo persona  (  J  ,F )

                            g.tipo_persona = Convert.ToString(workSheet.Cells[f, 27].Value).Substring(0, 1);


                            if (g.tipo_persona == "J")
                                g.razon_social = Convert.ToString(workSheet.Cells[f, 28].Value);
                            else
                            {
                                g.nombre_de_pila = Convert.ToString(workSheet.Cells[f, 29].Value);
                                g.primer_apellido = Convert.ToString(workSheet.Cells[f, 30].Value);
                            }

                            // Celda para el campo "telefono"   prefijo pais + numero ( S -  x(4) , x(12))
                            int celdaTelefono = 31;  // Índice de la columna para el teléfono
                            if (workSheet.Cells[f, celdaTelefono].Value != null)
                            {
                                string valorTelefono = Convert.ToString(workSheet.Cells[f, celdaTelefono].Value).Trim();

                                // Verificar si contiene solo un teléfono (sin comas)
                                if (!valorTelefono.Contains(","))
                                {
                                    string telefono = valorTelefono.Trim();

                                    // Verificar si el teléfono es válido: 6 a 12 dígitos, con o sin prefijo "+"
                                    if (System.Text.RegularExpressions.Regex.IsMatch(telefono, @"^(\d{6,12})$"))
                                    {
                                        g.telefono = telefono;  // Asignar el teléfono válido
                                        g.tlf_contacto = telefono;
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Teléfono: el número '{telefono}' no es válido. Debe tener entre 6 y 12 dígitos.");
                                        lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Teléfono: el campo contiene más de un número de teléfono.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Teléfono: el campo está vacío.");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //   old g.telefono = Convert.ToString(workSheet.Cells[f, 23].Value);


                            //   persona de contacto 
                            g.persona_contacto = Convert.ToString(workSheet.Cells[f, 32].Value);

                            //  indicatidor tipo direccion  F , s    -  dic_direccion_fiscal = Carga_Tabla_CNMC("cnmc_p_direccion_fiscal");
                            int celdaTipoDire = 33;  // Índice de la columna para "tipo dirección"

                            if (workSheet.Cells[f, celdaTipoDire].Value != null && Convert.ToString(workSheet.Cells[f, celdaTipoDire].Value).Trim().Length >= 1)
                            {
                                string valorTipoDire = Convert.ToString(workSheet.Cells[f, celdaTipoDire].Value).Substring(0, 1).ToUpper().Trim();

                                // Verificar si el valor está en el diccionario (clave o valor según tu necesidad)
                                if (cnmc.dic_direccion_fiscal.ContainsValue(valorTipoDire))
                                {
                                    g.indicador_tipo_direccion = valorTipoDire;  // Asignar el valor
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Indicador tipo dirección: el campo no contiene un valor válido.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Indicador tipo dirección: el campo es nulo o no tiene una longitud válida (mínimo 1 carácter).");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            // old      g.indicador_tipo_direccion = Convert.ToString(workSheet.Cells[f, 25].Value).Substring(0,1);


                            // DIRECCION 
                            g.pais = Convert.ToString(workSheet.Cells[f, 34].Value);


                            // privincia   - s - x(2) -  no tabla 
                            //  pero esta definida __    dic_provincias = Carga_Tabla_CNMC("cnmc_p_provincias");

                            int celdaProvincia = 35;       // Celda para el campo "provincia"
                            string valorProvincia = "00";
                            // Validar el campo "provincia"
                            if (workSheet.Cells[f, celdaProvincia].Value != null && Convert.ToString(workSheet.Cells[f, celdaProvincia].Value).Length >= 2)
                            {
                                valorProvincia = Convert.ToString(workSheet.Cells[f, celdaProvincia].Value).Substring(0, 2).ToUpper();

                                // Verificar si el valor de la provincia existe en el diccionario "cnmc_p_provincias"
                                if (cnmc.dic_provincias.ContainsValue(valorProvincia))
                                {
                                    g.provincia = valorProvincia; // Asignar el valor de la provincia
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " provincia: el campo provincia no contiene un valor válido.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + "provincia: el campo provincia es nulo o no tiene una longitud válida (2 caracteres).");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //   old   g.provincia = Convert.ToString(workSheet.Cells[f, 27].Value).Substring(0,2);

                            // municipio -Codigo INE compuesto por la concatenación de dos dígitos del código provincia (CPRO) más tres dígitos del código de municipio (CMUN)
                            // más un dígito de control (DC) adicional y opcional al final.
                            //   S - X(6)  ---  no tabla 
                            //      dic_municipios = Carga_Municipios();  ----  
                            int celdaMunicipio = 36;       // Celda para el campo "municipio"
                            if (workSheet.Cells[f, celdaMunicipio].Value != null && Convert.ToString(workSheet.Cells[f, celdaMunicipio].Value).Length >= 2)
                            {
                                //if (cnmc.dic_municipios.ContainsKey(valorProvincia+"|"+Convert.ToString(workSheet.Cells[f, celdaMunicipio].Value).ToLower()))
                                string cod_municipio;
                                if (cnmc.dic_municipios.TryGetValue(valorProvincia + "|" + Convert.ToString(workSheet.Cells[f, celdaMunicipio].Value).ToLower(), out cod_municipio))
                                {
                                    g.municipio = cod_municipio;
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " municipio: el campo tipo municipios no contiene un valor válido.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + " municipio: el campo municipio es nulo o no tiene una longitud válida X(2).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //  old  g.municipio = Convert.ToString(workSheet.Cells[f, 28].Value);

                            //                     codigo postal 
                            int celdaCP = 37;  // Índice de la columna para el Código Postal

                            if (workSheet.Cells[f, celdaCP].Value != null)
                            {
                                string valorCP = Convert.ToString(workSheet.Cells[f, celdaCP].Value).Trim();

                                if (int.TryParse(valorCP, out int codigoPostal))
                                {
                                    string codigoPostalStr = codigoPostal.ToString().PadLeft(5, '0');  // Completar con ceros a la izquierda si tiene menos de 5 dígitos

                                    if (codigoPostalStr.Length == 5)
                                    {
                                        g.codigo_postal = codigoPostalStr;
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Código Postal: debe contener exactamente 5 dígitos numéricos.");
                                        lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Código Postal: el valor ingresado no es numérico.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Código Postal: el campo está vacío.");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            // old   g.codigo_postal = Convert.ToString(workSheet.Cells[f, 29].Value).PadLeft(5, '0');

                            //  tipovia - -S- X(2)   -  Tabla 12
                            //         dic_via = Carga_Tabla_CNMC("cnmc_p_tipo_via");
                            int celdatipovia = 38;       // Celda para el campo "tipo de via "
                            // Validar el campo "tipo de via "
                            if (workSheet.Cells[f, celdatipovia].Value != null && Convert.ToString(workSheet.Cells[f, celdatipovia].Value).Length >= 2)
                            {
                                string valortipo_via = Convert.ToString(workSheet.Cells[f, celdatipovia].Value).Substring(0, 2).ToUpper();

                                // Verificar si el valor de la tipo via  en el diccionario "cnmc_p_tipo_via"
                                if (cnmc.dic_via.ContainsValue(valortipo_via))
                                {
                                    g.tipo_via = valortipo_via; // Asignar el valor de la tipo via 
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> tipo via : el campo tipo via  no contiene un valor válido.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> tipo via : el campo tipo via es nulo o no tiene una longitud válida (2 caracteres).");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //   old  g.tipo_via = Convert.ToString(workSheet.Cells[f, 30].Value).Substring(0,2);

                            //    nombre de via  ------31
                            var valorCelda = workSheet.Cells[f, 39].Value;

                            if (valorCelda != null)
                            {
                                string nombreVia = valorCelda.ToString().Trim(); // Elimina espacios innecesarios

                                if (nombreVia.Length > 30)
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " +
                                                  "Nombre de Vía: el texto excede los 30 caracteres y será truncado.");
                                    //lista_log.Add(System.Environment.NewLine);

                                    g.nombre_via = nombreVia.Substring(0, 30); // Corta a 30 caracteres
                                }
                                else
                                {
                                    g.nombre_via = nombreVia;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " +
                                              "Nombre de Vía: el campo está vacío o es nulo.");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //  old  g.nombre_via = Convert.ToString(workSheet.Cells[f, 31].Value);




                            //  numero  -  S -  X(5)
                            int celdaNumero = 40;  // Índice de la columna para "número"
                            if (workSheet.Cells[f, celdaNumero].Value != null)
                            {
                                string numero = Convert.ToString(workSheet.Cells[f, celdaNumero].Value).Trim();  // Elimina espacios adicionales

                                if (numero.Length > 5)
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Número: el texto excede los 5 caracteres y será truncado.");
                                    lista_log.Add(Environment.NewLine);

                                    g.numero = numero.Substring(0, 5);  // Trunca el valor a 5 caracteres
                                }
                                else if (numero.Length > 0)
                                {
                                    g.numero = numero;  // Asigna el valor tal cual si tiene entre 1 y 5 caracteres
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Número: el campo está vacío.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Número: el campo está nulo.");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //  old   g.numero = Convert.ToString(workSheet.Cells[f, 32].Value);

                            //  añadir   piso   del xls -
                            //    dic_piso = Carga_Tabla_CNMC("cnmc_p_piso")
                            //g.piso = Convert.ToString(workSheet.Cells[f, 32].Value);  //  Piso N x(3) Tabla 14
                            int celdapiso = 41;  // Celda para el campo "piso"
                            var valorCeldaPiso = workSheet.Cells[f, celdapiso]?.Value?.ToString()?.Trim().ToUpper();

                            if (!string.IsNullOrEmpty(valorCeldaPiso))
                            {
                                // Separar el valor si contiene un guion (-) y tomar la primera parte (número)
                                string[] partesPiso = valorCeldaPiso.Split('-');
                                string codigoPiso = partesPiso[0].Trim();

                                // Verificar si la primera parte es un número de hasta 3 dígitos
                                if (codigoPiso.Length <= 3 && int.TryParse(codigoPiso, out int numeroPiso))
                                {
                                    g.piso = numeroPiso.ToString("D3");  // Convertir a formato de 3 dígitos (001, 002, etc.)
                                }
                                else
                                {
                                    // Verificar si el valor abreviado está en el diccionario
                                    string codigoPisoAbreviado = codigoPiso.Length >= 2 ? codigoPiso.Substring(0, 2) : codigoPiso;

                                    if (cnmc.dic_piso.ContainsValue(codigoPisoAbreviado))
                                    {
                                        g.piso = codigoPisoAbreviado;  // Asignar el código abreviado directamente
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> El valor del campo 'piso' ({codigoPiso}) no es válido.");
                                        registro_valido = false;
                                    }
                                }
                            }
                            else
                            {
                                // Si la celda está vacía, mantener el valor vacío
                                g.piso = "";
                            }

                            //lista_log.Add(Environment.NewLine);

                            //  añadir  puerta   dic_puerta = Carga_Tabla_CNMC("cnmc_p_puerta");
                            // g.Puerta= Convert.ToString(workSheet.Cells[f, 32].Value); // Puerta N(3) Tabla 15
                            int celdaPuerta = 42;  // Celda para el campo "puerta"
                            var valorCeldaPuerta = workSheet.Cells[f, celdaPuerta]?.Value?.ToString()?.Trim().ToUpper();

                            if (!string.IsNullOrEmpty(valorCeldaPuerta))
                            {
                                // Separar el valor si contiene un guion (-) y tomar la primera parte (número)
                                string[] partesPuerta = valorCeldaPuerta.Split('-');
                                string codigoPuerta = partesPuerta[0].Trim();

                                // Verificar si la primera parte es un número de hasta 3 dígitos
                                if (codigoPuerta.Length <= 3 && int.TryParse(codigoPuerta, out int numeroPuerta))
                                {
                                    g.puerta = numeroPuerta.ToString("D3");  // Convertir a formato de 3 dígitos (001, 002, etc.)
                                }
                                else
                                {
                                    // Verificar si el valor abreviado está en el diccionario
                                    string codigoPuertaAbreviado = codigoPuerta.Length >= 2 ? codigoPuerta.Substring(0, 2) : codigoPuerta;

                                    if (cnmc.dic_puerta.ContainsValue(codigoPuertaAbreviado))
                                    {
                                        g.puerta = codigoPuertaAbreviado;  // Asignar el código abreviado directamente
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> El valor del campo 'puerta' ({codigoPuerta}) no es válido.");
                                        registro_valido = false;
                                    }
                                }
                            }
                            else
                            {
                                // Si la celda está vacía, mantener el valor vacío
                                g.puerta = "";
                            }

                            //lista_log.Add(Environment.NewLine);


                            #endregion

                            //--   control para la ofiina - Operacciones  -- 

                            g.entorno = Convert.ToString(workSheet.Cells[f, 43].Value).Trim();

                            //
                            g.tipo_cliente = Convert.ToString(workSheet.Cells[f, 44].Value).Trim();


                            //

                            #region RegistroDocumentoAD_A3 
                            //// TipoDocAportado   - s-  X(2)  tabla 61 - registro documento maxiom 50 ocurrencias 
                            ////  dic_documentacion = Carga_Tabla_CNMC("cnmc_p_tipo_documentacion");
                            //// DireccionUrl
                            c = 44;

                            try
                            {
                                for (int j = 1; j <= 6; j++)  // Bucle para verificar hasta 6 pares
                                {
                                    c++;
                                    Documentacion doc = new Documentacion();

                                    string tipoDoc = Convert.ToString(workSheet.Cells[f, c].Value)?.Trim();
                                    string urlDireccion = Convert.ToString(workSheet.Cells[f, c + 1].Value)?.Trim();

                                    // Si ambos valores son nulos o vacíos, continuar con el siguiente par
                                    if (string.IsNullOrEmpty(tipoDoc) && string.IsNullOrEmpty(urlDireccion))
                                    {
                                        c++;  // Avanzar dos columnas (porque estamos verificando pares TipoDoc, URL)
                                        continue;
                                    }

                                    // Obtener solo las 2 primeras posiciones de tipoDoc
                                    tipoDoc = !string.IsNullOrEmpty(tipoDoc) ? tipoDoc.Substring(0, Math.Min(2, tipoDoc.Length)) : "00";

                                    // Asignar valores al objeto Documentacion
                                    doc.tipo_doc_aportado = tipoDoc;
                                    doc.direccion_url = urlDireccion;

                                    // Agregar el objeto Documentacion a la lista
                                    g.lista_documentacion.Add(doc);

                                    c++;  // Avanzar la columna de la URL para la siguiente iteración
                                }
                            }
                            catch (Exception ex)
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {f} Columna: {c} --> Error inesperado: {ex.Message}");
                            }
                            //primera opcion
                            //c = 36;

                            //for (int j = 1; j <= 10; j++)
                            //{
                            //    c++;
                            //    Documentacion doc = new Documentacion();
                            //    if (Convert.ToString(workSheet.Cells[f, c].Value) != "" && Convert.ToString(workSheet.Cells[f, c + 1].Value) != "")
                            //    {
                            //        doc.tipo_doc_aportado = Convert.ToString(workSheet.Cells[f, c].Value);
                            //        doc.tipo_doc_aportado = doc.tipo_doc_aportado.Substring(0, doc.tipo_doc_aportado.IndexOf(' ')).PadLeft(2, '0');
                            //        doc.direccion_url = Convert.ToString(workSheet.Cells[f, c + 1].Value);
                            //        g.lista_documentacion.Add(doc);
                            //    }
                            //}

                            #endregion
                            if (registro_valido)
                            {
                                // MessageBox.Show( " registro valido -- ",  // se ha añadido mensaje ...informativo
                                //  "Extrasistemas.CargaExcel",
                                //    MessageBoxButtons.OK,
                                //      MessageBoxIcon.Information);
                                lista_datos_excel_extrasistemas.Add(g);
                                if (dic_total_registros.TryGetValue(g.hoja, out o))
                                    o.registros = o.registros + 1;
                            }

                        }
                    }

                    // C101 Alta Cambio Comercializadora Sin Cambios
                    else if (workSheet.Name == "ACC")
                    {
                        //NO PROCESAMOS ESTA PESTAÑA HASTA FINALIZAR IMPLEMENTACION
                        //FORZAMOS BREAK
                        //lista_log.Add("La pestaña ACC no se procesará, está deshabilitada temporalmente");
                        //continue;
                        //
                        f = 1; // Porque la primera fila es la cabecera
                        for (int i = 1; i < 1000000; i++)
                        {
                            registro_valido = true;
                            f++;
                            if (workSheet.Cells[f, 3].Value == null)
                                break;

                            acc_total++;

                            EndesaEntity.extrasistemas.Global g = new EndesaEntity.extrasistemas.Global();
                            g.hoja = workSheet.Name;
                            //CodigoREEEmpresaEmisora - Obligatorio X(4) - STRING NUMERICO
                            if (workSheet.Cells[f, 1].Value != null && Convert.ToString(workSheet.Cells[f, 1].Value).Length >= 4)
                            {
                                if (int.TryParse(Convert.ToString(workSheet.Cells[f, 1].Value).Substring(0, 4), out _))
                                    g.empresa_emisora = Convert.ToString(workSheet.Cells[f, 1].Value).Substring(0, 4);
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " Empresa_emisora: el campo código empresa emisora no tiene un formato válido.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " Empresa_emisora: el campo código empresa emisora es nulo o no tiene una longitud válida X(4).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //CodigoREEEmpresaDestino - Obligatorio X(4) - STRING NUMERICO
                            //g.distribuidora = Convert.ToString(workSheet.Cells[f, 2].Value);
                            if (workSheet.Cells[f, 2].Value != null && Convert.ToString(workSheet.Cells[f, 2].Value).Length >= 4)
                            {
                                if (int.TryParse(Convert.ToString(workSheet.Cells[f, 2].Value).Substring(0, 4), out _))
                                    g.distribuidora = Convert.ToString(workSheet.Cells[f, 2].Value).Substring(0, 4);
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " Distribuidora: el campo código empresa destino (distribuidora) no tiene un formato válido.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " Distribuidora: el campo código empresa destino (distribuidora) es nulo o no tiene una longitud válida X(4).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }

                            //CUPS - Obligatorio X(22)
                            //g.cups22 = Convert.ToString(workSheet.Cells[f, 3].Value);
                            if (workSheet.Cells[f, 3].Value != null && Convert.ToString(workSheet.Cells[f, 3].Value).Length >= 22)
                            {
                                g.cups22 = Convert.ToString(workSheet.Cells[f, 3].Value).Substring(0, 22);
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " CUPS: el campo CUPS es nulo o no tiene una longitud válida X(22).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //
                            // Variable para la celda de IndActivacion
                            int celdaIndActivacion = 4;
                            if (workSheet.Cells[f, celdaIndActivacion].Value != null && Convert.ToString(workSheet.Cells[f, celdaIndActivacion].Value).Length >= 1)
                            {
                                string indActivacion = Convert.ToString(workSheet.Cells[f, celdaIndActivacion].Value).Substring(0, 1).ToUpper();

                                if (cnmc.dic_indicativo_activacion.ContainsValue(indActivacion))
                                {
                                    g.ind_activacion = indActivacion;
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> IndActivacion: el campo IndActivacion no contiene un valor válido.");
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> IndActivacion: el campo IndActivacion es nulo o no tiene una longitud válida X(1).");
                                registro_valido = false;
                            }

                            if (g.ind_activacion == "F")
                            {
                                // Nos aseguramos que el campo contiene una fecha válida
                                string fechaAc = Convert.ToString(workSheet.Cells[f, 5].Value);

                                if (DateTime.TryParseExact(fechaAc, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out DateTime fechaActivacion))
                                {
                                    g.fecha_activacion = fechaActivacion;
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Fechaactivacion: el campo fecha_activación no tiene un campo fecha válido AAAA-MM-DD.");
                                    registro_valido = false;
                                }
                            }
                            //IndActivacion - Obligatorio X(1) [A,L,F]
                            // g.ind_activacion = Convert.ToString(workSheet.Cells[f, 4].Value).Substring(0, 1);
                            //if (workSheet.Cells[f, 4].Value != null && Convert.ToString(workSheet.Cells[f, 4].Value).Length >= 1)
                            //{
                            //    if (cnmc.dic_indicativo_activacion.ContainsValue(Convert.ToString(workSheet.Cells[f, 4].Value).Substring(0, 1).ToUpper()))
                            //    {
                            //        g.ind_activacion = Convert.ToString(workSheet.Cells[f, 4].Value).Substring(0, 1);
                            //    }
                            //    else
                            //    {
                            //        lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " IndActivacion: el campo IndActivacion no contiene un valor válido.");
                            //        //lista_log.Add(System.Environment.NewLine);
                            //        registro_valido = false;
                            //    }
                            //}
                            //else
                            //{
                            //    lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + " IndActivacion: el campo IndActivacion es nulo o no tiene una longitud válida X(1).");
                            //    //lista_log.Add(System.Environment.NewLine);
                            //    registro_valido = false;
                            //}


                            ////FechaPrevistaAccion - Condicionada a valor IndActivacion == F - Fecha(AAAA-MM-DD)
                            //if (g.ind_activacion == "F")
                            //{

                            //    // Nos aseguramos que el campo contiene una fecha válida
                            //    fecha = Convert.ToString(workSheet.Cells[f, 5].Value);

                            //    if (DateTime.TryParse(fecha, out _))
                            //    {
                            //        g.fecha_activacion = Convert.ToDateTime(fecha);

                            //    }
                            //    else
                            //    {
                            //        lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + " FechaPrevistaAccion: el campo fecha_activación no tiene un campo fecha válido AAAA-MM-DD.");
                            //        //lista_log.Add(System.Environment.NewLine);
                            //        registro_valido = false;
                            //    }

                            //}


                            //ContratacionIncondicionalPS - Obligatorio X(1) [S,N]
                            //g.contratacion_incondicional_ps = Convert.ToString(workSheet.Cells[f, 6].Value).Substring(0, 1);
                            if (workSheet.Cells[f, 6].Value != null && Convert.ToString(workSheet.Cells[f, 6].Value).Length >= 1)
                            {
                                if (cnmc.dic_indicativo_sino.ContainsValue(Convert.ToString(workSheet.Cells[f, 6].Value).Substring(0, 1).ToUpper()))
                                {
                                    g.contratacion_incondicional_ps = Convert.ToString(workSheet.Cells[f, 6].Value).Substring(0, 1).ToUpper();
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " ContratacionIncondicionalPS: el campo ContratacionIncondicionalPS no contiene un valor válido.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + " ContratacionIncondicionalPS: el campo ContratacionIncondicionalPS es nulo o no tiene una longitud válida X(1).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }

                            //ContratacionIncondicionalBS - Obligatorio X(1) [S,N]
                            //Siempre a 'N'
                            g.contratacion_incondicional_bs = "N";

                            // tipo indentificador  buscar en diccionario 
                            if (workSheet.Cells[f, 8].Value != null)
                            {
                                // Normalizar el valor de la celda
                                string descripcion = Convert.ToString(workSheet.Cells[f, 8].Value)?.Trim().ToUpperInvariant();

                                if (!string.IsNullOrEmpty(descripcion))
                                {
                                    // Normalizar claves del diccionario eliminando espacios y asegurando mayúsculas
                                    var dic_normalizado = cnmc.dic_identificador
                                        .ToDictionary(kv => kv.Key.Trim().ToUpperInvariant(), kv => kv.Value);

                                    // Buscar la clave en el diccionario
                                    if (!dic_normalizado.TryGetValue(descripcion, out string valorEncontrado))
                                    {
                                        // Intento adicional: Buscar con coincidencias parciales si las claves son listas
                                        foreach (var kvp in dic_normalizado)
                                        {
                                            var clavesDiccionario = kvp.Key.Split(new[] { ',', '-', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                                           .Select(s => s.Trim().ToUpperInvariant());

                                            var clavesEntrada = descripcion.Split(new[] { ',', '-', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                                           .Select(s => s.Trim().ToUpperInvariant());

                                            if (clavesDiccionario.Intersect(clavesEntrada).Any()) // Coincidencia parcial
                                            {
                                                valorEncontrado = kvp.Value;
                                                break;
                                            }
                                        }
                                    }

                                    // Verificar si encontramos el valor  - tipo de indetificador 
                                    if (!string.IsNullOrEmpty(valorEncontrado))
                                    {
                                        if (valorEncontrado.Length >= 2)
                                        {
                                            g.tipo_identificador = valorEncontrado.Substring(valorEncontrado.Length - 2).ToUpper();
                                        }
                                        else
                                        {
                                            lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El valor asociado a '{descripcion}' en el diccionario es inválido ({valorEncontrado}).");
                                            //lista_log.Add(Environment.NewLine);
                                            registro_valido = false;
                                        }
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: La descripción '{descripcion}' no se encontró en el diccionario.");
                                        //lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El valor en la celda es inválido (vacío o solo espacios).");
                                    //lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El campo es nulo.");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //TipoIdentificador - Obligatorio X(2) [Tabla 6]
                            //g.tipo_identificador = Convert.ToString(workSheet.Cells[f, 8].Value);
                            //if (workSheet.Cells[f, 8].Value != null && Convert.ToString(workSheet.Cells[f, 8].Value).Length >= 2)
                            //{
                            //    // irh  - fila 8 / H : tipo_identificador  ( nif , nie, etc ...)se cambia por el key ..
                            //    //if (cnmc.dic_identificador.ContainsValue(Convert.ToString(workSheet.Cells[f, 8].Value).Substring(0, 2).ToUpper()))

                            //    if (cnmc.dic_identificador.ContainsKey(Convert.ToString(workSheet.Cells[f, 8].Value).ToUpper()))
                            //    {
                            //        //g.tipo_identificador = Convert.ToString(workSheet.Cells[f, 8].Value).Substring(0, 2).ToUpper();
                            //        g.tipo_identificador = Convert.ToString(workSheet.Cells[f, 8].Value).ToUpper();
                            //    }
                            //    else
                            //    {
                            //        lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " TipoIdentificador: el campo Tipo_indentificador no contiene un valor válido.");
                            //        //lista_log.Add(System.Environment.NewLine);
                            //        registro_valido = false;
                            //    }
                            //}
                            //else
                            //{
                            //    lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + " TipoIdentificador: el campo Tipo_indentificador es nulo o no tiene una longitud válida X(2).");
                            //    //lista_log.Add(System.Environment.NewLine);
                            //    registro_valido = false;
                            //}

                            //Identificador - Obligatorio X(14) Longitud según TipoIdentificador, como máximo 14
                            g.n_identificador = Convert.ToString(workSheet.Cells[f, 9].Value);



                            //TipoPersona - NO Obligatorio X(1) [F, J]
                            g.tipo_persona = Convert.ToString(workSheet.Cells[f, 10].Value).Substring(0, 1);

                            if (g.tipo_persona == "J")
                                g.razon_social = Convert.ToString(workSheet.Cells[f, 11].Value);
                            else
                            {
                                g.nombre_de_pila = Convert.ToString(workSheet.Cells[f, 12].Value);
                                g.primer_apellido = Convert.ToString(workSheet.Cells[f, 13].Value);
                            }

                            g.telefono = Convert.ToString(workSheet.Cells[f, 14].Value);
                            g.entorno = Convert.ToString(workSheet.Cells[f, 15].Value);
                            g.tipo_cliente = Convert.ToString(workSheet.Cells[f, 16].Value);



                            if (Convert.ToString(workSheet.Cells[f, 17].Value) != "")
                                g.observaciones = Convert.ToString(workSheet.Cells[f, 17].Value);

                            if (registro_valido)
                            {
                                lista_datos_excel_extrasistemas.Add(g);
                                if (dic_total_registros.TryGetValue(g.hoja, out o))
                                    o.registros = o.registros + 1;
                            }
                        }

                    }
                    // C201 Alta Cambio Comercializadora Con Cambios
                    else if (workSheet.Name == "ACC + CAMBIOS")  // xls
                    {

                        //NO PROCESAMOS ESTA PESTAÑA HASTA FINALIZAR IMPLEMENTACION
                        //FORZAMOS BREAK
                       //lista_log.Add("La pestaña ACC+ CAMBIOS no se procesará, está deshabilitada temporalmente");
                       // continue;
                        //

                      f = 2; // Porque la primera fila es la cabecera
                      for (int i = 1; i < 1000000; i++)
                      {

                            registro_valido = true;
                            f++;
                            if (workSheet.Cells[f, 3].Value == null)
                                break;

                            acc_cambios_total++;

                            EndesaEntity.extrasistemas.Global g = new EndesaEntity.extrasistemas.Global();
                            g.hoja = workSheet.Name;
                            //
                            //  CodigoREEEmpresaEmisora -  1
                            if (workSheet.Cells[f, 1].Value != null && Convert.ToString(workSheet.Cells[f, 1].Value).Length >= 4)
                            {
                                if (int.TryParse(Convert.ToString(workSheet.Cells[f, 1].Value).Substring(0, 4), out _))

                                    g.empresa_emisora = Convert.ToString(workSheet.Cells[f, 1].Value).Substring(0, 4);
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " Empresa_emisora: el campo código empresa emisora no tiene un formato válido.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " Empresa_emisora: el campo código empresa emisora es nulo o no tiene una longitud válida X(4).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //

                            int celdaDistribuidora = 2;  // Variable para la celda de Distribuidora
                            if (workSheet.Cells[f, celdaDistribuidora].Value != null && Convert.ToString(workSheet.Cells[f, celdaDistribuidora].Value).Length >= 4)
                            {
                                if (int.TryParse(Convert.ToString(workSheet.Cells[f, celdaDistribuidora].Value).Substring(0, 4), out _))
                                {
                                    g.distribuidora = Convert.ToString(workSheet.Cells[f, celdaDistribuidora].Value).Substring(0, 4);
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Distribuidora: el campo código empresa destino (distribuidora) no tiene un formato válido.");
                                    //lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Distribuidora: el campo código empresa destino (distribuidora) es nulo o no tiene una longitud válida X(4).");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //

                            int celdaCUPS = 3;   // Variable para la celda de CUPS
                            if (workSheet.Cells[f, celdaCUPS].Value != null && Convert.ToString(workSheet.Cells[f, celdaCUPS].Value).Length >= 22)
                            {
                                g.cups22 = Convert.ToString(workSheet.Cells[f, celdaCUPS].Value).Substring(0, 22);
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> CUPS: el campo CUPS es nulo o no tiene una longitud válida X(22).");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }


                            // tipo modificacion 4
                            int celdatipomodificacion = 4;
                            if (workSheet.Cells[f, celdatipomodificacion].Value != null && Convert.ToString(workSheet.Cells[f, celdatipomodificacion].Value).Length >= 1)
                            {
                                g.tipo_modificacion = Convert.ToString(workSheet.Cells[f, celdatipomodificacion].Value).Substring(0, 1);
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Tipo_modificacion: el campo tipo de modificación es nulo o no tiene una longitud válida X(1).");
                                registro_valido = false;
                            }

                            //  tipo solicitud administrativa 5
                            int celdatiposolicitudadministrativa = 5;
                            if (g.tipo_modificacion == "A" || g.tipo_modificacion == "S")
                            {
                                if (workSheet.Cells[f, celdatiposolicitudadministrativa].Value != null && Convert.ToString(workSheet.Cells[f, celdatiposolicitudadministrativa].Value).Length >= 1)
                                {
                                    g.tipo_solicitud_administrativa = Convert.ToString(workSheet.Cells[f, celdatiposolicitudadministrativa].Value).Substring(0, 1);
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Tipo_solicitud_administrativa: el campo es requerido para modificaciones de tipo 'A' o 'S' y es nulo o tiene una longitud inválida X(1).");
                                    registro_valido = false;
                                }
                            }
                            //
                            int celdaCNAE = 6;   // Variable para la celda de CNAE
                            if (workSheet.Cells[f, celdaCNAE].Value != null && Convert.ToString(workSheet.Cells[f, celdaCNAE].Value).Length >= 4)
                            {
                                g.cnae = Convert.ToString(workSheet.Cells[f, celdaCNAE].Value).Substring(0, 4);
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> CUPS: {g.cups22}  CNAE: el campo CNAE es nulo o no tiene una longitud válida X(4).");
                                //lista_log.Add($"Hoja: {g.hoja} --> CUPS: {g.cups22} El campo 'persona de contacto' es obligatorio y está vacío.")
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //
                            //IndActivacion ---------7
                            int celdaIndActivacion = 7;

                            if (workSheet.Cells[f, celdaIndActivacion].Value != null && Convert.ToString(workSheet.Cells[f, celdaIndActivacion].Value).Length >= 1)
                            {
                                string valorIndActivacion = Convert.ToString(workSheet.Cells[f, celdaIndActivacion].Value).Substring(0, 1).ToUpper();

                                if (cnmc.dic_indicativo_activacion.ContainsValue(valorIndActivacion))
                                {
                                    g.ind_activacion = valorIndActivacion;
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> CUPS: {g.cups22} IndActivacion: el campo IndActivacion no contiene un valor válido.");
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> CUPS: {g.cups22} IndActivacion: el campo IndActivacion es nulo o no tiene una longitud válida X(1).");
                                registro_valido = false;
                            }
                            //
                            int celdaFechaActivacion = 8;  // Variable para la celda de fecha_activacion = FechaPrevistaAccion
                            if (g.ind_activacion == "F")
                            {
                                var valorFecha = workSheet.Cells[f, celdaFechaActivacion].Value;

                                if (valorFecha != null)
                                {
                                    string fechaStr = Convert.ToString(valorFecha).Trim();

                                    if (DateTime.TryParse(fechaStr, out DateTime fechaParsed))
                                    {
                                        // Verifica que el formato sea exactamente AAAA-MM-DD
                                        if (fechaParsed.ToString("yyyy-MM-dd") == fechaStr)
                                        {
                                            g.fecha_activacion = fechaParsed;
                                        }
                                        else
                                        {
                                            lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> CUPS: {g.cups22}  el campo fecha_activación no tiene el formato AAAA-MM-DD.");
                                            //lista_log.Add(Environment.NewLine);
                                            registro_valido = false;
                                        }
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> CUPS: {g.cups22}  el campo fecha_activación no es una fecha válida.");
                                        //lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} -->  CUPS: {g.cups22}  el campo fecha_activación está vacío.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            //
                            //ContratacionIncondicionalPS - 9

                            if (workSheet.Cells[f, 9].Value != null && Convert.ToString(workSheet.Cells[f, 9].Value).Length >= 1)
                            {
                                if (cnmc.dic_indicativo_sino.ContainsValue(Convert.ToString(workSheet.Cells[f, 9].Value).Substring(0, 1).ToUpper()))
                                {
                                    g.contratacion_incondicional_ps = Convert.ToString(workSheet.Cells[f, 9].Value).Substring(0, 1).ToUpper();
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " ContratacionIncondicionalPS: el campo ContratacionIncondicionalPS no contiene un valor válido.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + " ContratacionIncondicionalPS: el campo ContratacionIncondicionalPS es nulo o no tiene una longitud válida X(1).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //

                            //ContratacionIncondicionalBS - Obligatorio  10  --Siempre a 'N'

                            g.contratacion_incondicional_bs = "N";

                            //tipo_contrato    11
                            int celda11 = 11;

                            if (workSheet.Cells[f, celda11].Value != null && Convert.ToString(workSheet.Cells[f, celda11].Value).Length >= 2)
                            {
                                string tipoContrato = Convert.ToString(workSheet.Cells[f, celda11].Value).Substring(0, 2).ToUpper();

                                if (cnmc.dic_contrato_atr.ContainsValue(tipoContrato))
                                {
                                    g.tipo_contrato_atr = tipoContrato;
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> tipo_contrato_at: el campo tipo_contrato_at no contiene un valor válido.");
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> Fila: " + i + " tipo_contrato_at: el campo tipo_contrato_at es nulo o no tiene una longitud válida (X2).");
                                registro_valido = false;
                            }
                            // 

                            //FECHA DE FINALIZACION  12
                            int celda12 = 12;
                            string[] contratosRequierenFecha = { "02", "03", "09" };

                            // Solo validar si el tipo de contrato requiere fecha
                            if (contratosRequierenFecha.Contains(g.tipo_contrato_atr))
                            {
                                string fecha12 = workSheet.Cells[f, celda12]?.Value?.ToString()?.Trim();

                                if (!string.IsNullOrEmpty(fecha12))
                                {
                                    if (DateTime.TryParseExact(fecha12, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fechaFinalizacion))
                                    {
                                        g.fecha_finalizacion = fechaFinalizacion;
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} CUPS: {g.cups22} ➝ Fecha Finalización: '{fecha12}' no es válida (debe ser AAAA-MM-DD).");
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} CUPS: {g.cups22} ➝ La Fecha Finalización es obligatoria para el tipo de contrato '{g.tipo_contrato_atr}'. Rechazo '08'.");
                                    registro_valido = false;
                                }
                            }

                            // tipoautoconsumo   13
                            int celda13 = 13;
                            string valorCelda13 = workSheet.Cells[f, celda13]?.Value?.ToString()?.Trim();

                            if (!string.IsNullOrEmpty(valorCelda13) && valorCelda13.Length >= 2)
                            {
                                string tipoAutoconsumo = valorCelda13.Substring(0, 2).ToUpper();

                                // Validar contra el diccionario (clave es la descripción)
                                if (cnmc.dic_autoconsumo.ContainsValue(tipoAutoconsumo))
                                {
                                    g.tipo_autoconsumo = tipoAutoconsumo;
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> tipo_autoconsumo: el campo no contiene un valor válido.");
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} tipo_autoconsumo: el campo es nulo o no tiene una longitud válida (mínimo 2 caracteres).");
                                registro_valido = false;
                            }
                            //tipocups  14 
                            // Verificar si se deben informar los campos (obligatorio si TipoAutoconsumo ≠ "00" y ≠ "0C")
                            bool debeInformarCampos = g.tipo_autoconsumo != "00" && g.tipo_autoconsumo != "0C";

                            if (debeInformarCampos)
                            {
                                int columnaTipoCUPS = 14;
                                var valorCelda = workSheet.Cells[f, columnaTipoCUPS].Value;
                                string valorTexto = Convert.ToString(valorCelda)?.Trim();

                                if (!string.IsNullOrEmpty(valorTexto) && valorTexto.Length >= 2)
                                {
                                    g.tipocups = valorTexto.Substring(0, 2).ToUpper();
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} | Error: el campo 'tipocups' es obligatorio cuando 'tipo_autoconsumo' es distinto de '00' y '0C', pero está vacío o no tiene longitud suficiente.");
                                    registro_valido = false;
                                }
                            }
                            //cau
                            int celdaCau = 15;
                            var valorCeldaCau = workSheet.Cells[f, celdaCau].Value;

                            if (valorCeldaCau != null && Convert.ToString(valorCeldaCau).Length == 26)
                            {
                                // Asignar el valor directamente
                                g.cau = Convert.ToString(valorCeldaCau);
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} cau: el campo cau es nulo o no tiene una longitud válida.");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //
                            int celdaTiposubseccion = 16;
                            var valorCeldaTiposubseccion = workSheet.Cells[f, celdaTiposubseccion].Value;

                            if (valorCeldaTiposubseccion != null && Convert.ToString(valorCeldaTiposubseccion).Length >= 2)
                            {
                                // Se obtiene el valor de las primeras dos letras en mayúsculas
                                string tiposubseccionValor = Convert.ToString(valorCeldaTiposubseccion).Substring(0, 2).ToUpper();

                                g.tipo_subseccion = tiposubseccionValor;
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} tiposubseccion: el campo tiposubseccion es nulo o no tiene una longitud válida.");
                                registro_valido = false;
                            }
                            //
                            int celdaColectivo = 17;
                            var valorCeldaColectivo = workSheet.Cells[f, celdaColectivo].Value;

                            if (valorCeldaColectivo != null && (Convert.ToString(valorCeldaColectivo) == "S" || Convert.ToString(valorCeldaColectivo) == "N"))
                            {
                                g.colectivo = Convert.ToString(valorCeldaColectivo);
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} colectivo: el campo colectivo es nulo o no tiene un valor válido.");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //
                            int celdaPotInstaladaGen = 18;

                            var valorCeldaPotInstaladaGen = workSheet.Cells[f, celdaPotInstaladaGen].Value;

                            if (valorCeldaPotInstaladaGen != null && long.TryParse(Convert.ToString(valorCeldaPotInstaladaGen), out long potInstaladaGenValor) && Convert.ToString(valorCeldaPotInstaladaGen).Length <= 14)
                            {
                                // Asignar el valor directamente
                                g.potinstaladagen = (int)potInstaladaGenValor;
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} potInstaladaGen: el campo potInstaladaGen es nulo, no es numérico o excede las 14 posiciones.");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //
                            //
                            int celdaTipoInstalacion = 19;
                            var valorCeldaTipoInstalacion = workSheet.Cells[f, celdaTipoInstalacion].Value;

                            if (valorCeldaTipoInstalacion != null && Convert.ToString(valorCeldaTipoInstalacion).Length >= 2)
                            {
                                string tipoInstalacionValor = Convert.ToString(valorCeldaTipoInstalacion).Substring(0, 2).ToUpper();

                                if (g.tipo_autoconsumo == "11")
                                {
                                    // Si TipoAutoconsumo es "11", solo se permite "01" o "02"
                                    if (tipoInstalacionValor == "01" || tipoInstalacionValor == "02")
                                    {
                                        g.tipoinstalacion = tipoInstalacionValor;
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} TipoInstalacion: cuando TipoAutoconsumo es '11', solo puede ser '01' o '02'. Se recibió '{tipoInstalacionValor}'.");
                                        //lista_log.Add(System.Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    // Si TipoAutoconsumo NO es "11", cualquier valor informado en TipoInstalacion es válido
                                    g.tipoinstalacion = tipoInstalacionValor;
                                }
                            }
                            else
                            {
                                // Si TipoAutoconsumo es "11", TipoInstalacion es obligatorio
                                if (g.tipo_autoconsumo == "11")
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} TipoInstalacion: cuando TipoAutoconsumo es '11', el campo TipoInstalacion es obligatorio ('01' o '02').");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                                else
                                {
                                    // Si el TipoAutoconsumo NO es "11", TipoInstalacion puede estar vacío (nulo)
                                    g.tipoinstalacion = null;
                                }
                            }
                            //
                            int celdaSsaa = 20;
                            var valorCeldaSsaa = workSheet.Cells[f, celdaSsaa].Value;

                            if (valorCeldaSsaa != null && (Convert.ToString(valorCeldaSsaa) == "S" || Convert.ToString(valorCeldaSsaa) == "N"))
                            {
                                // Asignar el valor directamente
                                g.ssaa = Convert.ToString(valorCeldaSsaa);
                            }
                            else if (g.tipocups == "01")
                            {
                                lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} ssaa: el campo ssaa es obligatorio y debe contener 'S' o 'N' cuando TipoCups es '01'.");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} ssaa: el campo ssaa es nulo o no tiene un valor válido.");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //
                            int celdaUnicoContrato = 21;
                            var valorCeldaUnicoContrato = workSheet.Cells[f, celdaUnicoContrato].Value;
                            if (g.ssaa == "S")
                            {
                                if (valorCeldaUnicoContrato != null && (Convert.ToString(valorCeldaUnicoContrato) == "S" || Convert.ToString(valorCeldaUnicoContrato) == "N"))
                                {

                                    g.unicocontrato = Convert.ToString(valorCeldaUnicoContrato);
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} unicocontrato: el campo unicocontrato es obligatorio y debe contener 'S' o 'N' cuando SSAA es 'S'.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                // Asignar el valor directamente si SSAA no es 'S'
                                g.unicocontrato = Convert.ToString(valorCeldaUnicoContrato);
                            }

                            //SOLICITUD MODIFICACION TENSION   
                            int celda22 = 22;
                            string valorCelda22 = workSheet.Cells[f, celda22]?.Value?.ToString()?.Trim();

                            if (!string.IsNullOrEmpty(valorCelda22))
                            {
                                // Tomamos el primer carácter en mayúscula
                                char inicial = char.ToUpper(valorCelda22[0]);

                                if (inicial == 'S' || inicial == 'N')
                                {
                                    g.solicitud_modificacion_tension = inicial.ToString();
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} ➝ Fila: {i} ➝ solicitud_modificacion_tension: valor no válido '{valorCelda22}'. Debe comenzar con 'S' o 'N'.");
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} ➝ Fila: {i} ➝ solicitud_modificacion_tension: el campo está vacío.");
                                registro_valido = false;
                            }

                            // Tension solicitada --      tabla  -   mysql cnmc_p_tensiones
                            int celdatensionsolicitada = 23;

                            if (g.solicitud_modificacion_tension == "S")
                            {
                                if (workSheet.Cells[f, celdatensionsolicitada].Value != null && Convert.ToString(workSheet.Cells[f, celdatensionsolicitada].Value).Length >= 2)
                                {
                                    string valorTension = Convert.ToString(workSheet.Cells[f, celdatensionsolicitada].Value).Substring(0, 2);

                                    if (cnmc.dic_tensiones.ContainsValue(valorTension))
                                    {
                                        g.tension_solicitada = valorTension;
                                    }
                                    else
                                    {
                                        lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Tension_solicitada: el valor '" + valorTension + "' no se encuentra en el diccionario de tensiones válido.");
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Tension_solicitada: campo requerido porque solicitud_modificacion_tension es 'S', pero es nulo o no tiene una longitud válida X(2).");
                                    registro_valido = false;
                                }
                            }
                            //
                            // Tarifa -  24
                            var celdaTarifa = workSheet.Cells[f, 24].Value;
                            if (celdaTarifa != null && Convert.ToString(celdaTarifa).Length >= 3)
                            {
                                string tarifaStr = Convert.ToString(celdaTarifa).Substring(0, 3).ToUpper();
                                // Ojo: la clave es la descripción (ContainsValue)
                                if (cnmc.dic_tarifa_atr.ContainsValue(tarifaStr))
                                {
                                    g.tarifa = tarifaStr;
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> tarifa atr: el campo tarifa atr no contiene un valor válido.");
                                    lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> Fila: " + i + " tarifa atr: el campo tarifa atr es nulo o no tiene una longitud válida (X3).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }

                            // Validación de las potencias contratadas en vatios (máximo 6)
                            int potenciasInformadas = 0;
                            bool esTarifa20TD = (g.tarifa == "2.0TD");

                            for (int pot = 0; pot < 6; pot++)
                            {
                                var celdaPotencia = workSheet.Cells[f, 25 + pot].Value;
                                string potenciaStr = Convert.ToString(celdaPotencia);

                                if (!string.IsNullOrWhiteSpace(potenciaStr) && int.TryParse(potenciaStr, out int potencia))
                                {
                                    if (potencia >= 0)
                                    {
                                        g.potencias[pot] = potencia;
                                        potenciasInformadas++;
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Potencia {pot + 1}: el valor debe ser un número positivo.");
                                        //lista_log.Add(Environment.NewLine);  CUPS: {g.cups22}
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Potencia {pot + 1}: el valor no es un número válido o está en blanco.");
                                    //lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }

                            //// Validación mínima: al menos 3 potencias deben estar informadas
                            //if (potenciasInformadas < 3)
                            //{
                            //    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Deben estar informadas al menos 3 potencias.");
                            //    lista_log.Add(Environment.NewLine);
                            //    registro_valido = false;
                            //}

                            // Validación de orden creciente de potencias (solo si no es tarifa 2.0TD)
                            if (!esTarifa20TD)
                            {
                                for (int pot = 1; pot < 6; pot++)
                                {
                                    // Si alguna potencia no está informada, se salta la validación de orden
                                    if (g.potencias[pot] == 0 || g.potencias[pot - 1] == 0) continue;

                                    if (g.potencias[pot] < g.potencias[pot - 1])
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Las potencias deben ser crecientes. Potencia {pot + 1} ({g.potencias[pot]}) es menor que Potencia {pot} ({g.potencias[pot - 1]}).");
                                        lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                        break;
                                    }
                                }
                            }
                            //
                            // ModoControlPotencia --- 31
                            var celdaModoControl = workSheet.Cells[f, 31].Value;
                            if (celdaModoControl != null && Convert.ToString(celdaModoControl).Length >= 1)
                            {
                                string modoControlStr = Convert.ToString(celdaModoControl).Substring(0, 1).ToUpper();

                                g.modo_control_potencia = modoControlStr;
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} CUPS: {g.cups22} --> El campo modo_control_potencia está vacío o no tiene una longitud válida.");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }

                            // Persona de contacto 32
                            var personaContactoRaw = Convert.ToString(workSheet.Cells[f, 32].Value);
                            if (!string.IsNullOrWhiteSpace(personaContactoRaw))
                            {
                                var personaContacto = personaContactoRaw.Substring(0, Math.Min(personaContactoRaw.Length, 150));
                                g.persona_contacto = personaContacto;
                                g.contacto = personaContacto;
                            }
                            else
                            {
                                registro_valido = false;
                                lista_log.Add($"Hoja: {g.hoja} --> CUPS: {g.cups22} El campo 'persona de contacto' es obligatorio y está vacío.");
                            }
                            // telefono de contacto 33
                            var telefonoRaw = Convert.ToString(workSheet.Cells[f, 33].Value);
                            // Validar que el teléfono no esté vacío y tenga entre 6 y 12 dígitos numéricos
                            if (!string.IsNullOrEmpty(telefonoRaw))
                            {
                                // Eliminar espacios y caracteres no numéricos si se desea (opcional)
                                var telefono = new string(telefonoRaw.Where(char.IsDigit).ToArray());

                                if (telefono.Length >= 6 && telefono.Length <= 12)
                                {
                                    g.telefono = telefono;
                                    g.tlf_contracto = telefono;
                                }
                                else
                                {
                                    registro_valido = false;
                                    lista_log.Add($"Hoja: {g.hoja} --> CUPS: {g.cups22} El campo teléfono debe tener entre 6 y 12 dígitos.");
                                }
                            }
                            else
                            {
                                registro_valido = false;
                                lista_log.Add($"Hoja: {g.hoja} --> CUPS: {g.cups22} El campo teléfono no tiene un valor válido.");
                            }

                            //Tipo_indentificador 34
                            var celdaIdentificador = workSheet.Cells[f, 34].Value;
                            if (celdaIdentificador != null)
                            {
                                // Normalizar el valor de la celda
                                string descripcion = Convert.ToString(celdaIdentificador)?.Trim().ToUpperInvariant();

                                if (!string.IsNullOrEmpty(descripcion))
                                {
                                    // Normalizar claves del diccionario eliminando espacios y asegurando mayúsculas
                                    var dic_normalizado = cnmc.dic_identificador
                                        .ToDictionary(kv => kv.Key.Trim().ToUpperInvariant(), kv => kv.Value);

                                    // Buscar coincidencia directa
                                    if (!dic_normalizado.TryGetValue(descripcion, out string valorEncontrado))
                                    {
                                        // Intento adicional: Buscar con coincidencias parciales si las claves son listas
                                        foreach (var kvp in dic_normalizado)
                                        {
                                            var clavesDiccionario = kvp.Key.Split(new[] { ',', '-', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                                           .Select(s => s.Trim().ToUpperInvariant());

                                            var clavesEntrada = descripcion.Split(new[] { ',', '-', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                                           .Select(s => s.Trim().ToUpperInvariant());

                                            if (clavesDiccionario.Intersect(clavesEntrada).Any()) // Coincidencia parcial
                                            {
                                                valorEncontrado = kvp.Value;
                                                break;
                                            }
                                        }
                                    }

                                    // Verificar si encontramos el valor
                                    if (!string.IsNullOrEmpty(valorEncontrado))
                                    {
                                        if (valorEncontrado.Length >= 2)
                                        {
                                            g.tipo_identificador = valorEncontrado.Substring(valorEncontrado.Length - 2).ToUpper();
                                        }
                                        else
                                        {
                                            lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El valor asociado a '{descripcion}' en el diccionario es inválido ({valorEncontrado}).");
                                            lista_log.Add(Environment.NewLine);
                                            registro_valido = false;
                                        }
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: La descripción '{descripcion}' no se encontró en el diccionario.");
                                        lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El valor en la celda es inválido (vacío o solo espacios).");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El campo es nulo.");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //

                            //N_Identificador 35
                            var celdaIdentifica = workSheet.Cells[f, 35].Value;

                            if (celdaIdentifica != null && Convert.ToString(celdaIdentifica).Trim().Length >= 1)
                            {
                                string valorIdentificador = Convert.ToString(celdaIdentifica).Trim();

                                // Validación de longitud máxima de 14 caracteres
                                if (valorIdentificador.Length <= 14)
                                {
                                    g.n_identificador = valorIdentificador.ToUpper();
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> n_identificador: El valor excede la longitud máxima permitida (14 caracteres).");
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> n_identificador: El campo es nulo o está vacío.");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }

                            //TipoPersona  36
                            var tipoPersona = Convert.ToString(workSheet.Cells[f, 36].Value);
                            // Comprobar si el tipo de persona no está vacío
                            if (!string.IsNullOrEmpty(tipoPersona))
                            {
                                g.tipo_persona = tipoPersona.Substring(0, 1);

                                if (g.tipo_persona == "J")
                                {
                                    var razonSocial = Convert.ToString(workSheet.Cells[f, 37].Value);
                                    g.razon_social = razonSocial?.Substring(0, Math.Min(razonSocial.Length, 50));
                                }
                                else
                                {
                                    var nombre = Convert.ToString(workSheet.Cells[f, 38].Value);
                                    var apellido = Convert.ToString(workSheet.Cells[f, 39].Value);

                                    g.nombre_de_pila = nombre?.Substring(0, Math.Min(nombre.Length, 50));
                                    g.primer_apellido = apellido?.Substring(0, Math.Min(apellido.Length, 50));
                                }
                            }

                            //IndicadorTipoDireccion  40
                            var indicadorRaw = Convert.ToString(workSheet.Cells[f, 40].Value);
                            if (!string.IsNullOrWhiteSpace(indicadorRaw))
                            {
                                g.indicador_tipo_direccion = indicadorRaw.Substring(0, 1);
                            }
                            else
                            {
                                registro_valido = false;
                                lista_log.Add($"Hoja: {g.hoja} --> CUPS: {g.cups22} El campo 'IndicadorTipoDireccion' no tiene un valor válido.");
                            }
                            // Direccion  Es obligatorio si el Indicador es "F" 

                            if (g.indicador_tipo_direccion == "F")
                            {

                                // pais 
                                string pais = Convert.ToString(workSheet.Cells[f, 41].Value)?.Trim();   // Obtener valor de la celda y convertirlo a string eliminando espacios
                                if (!string.IsNullOrEmpty(pais))
                                {
                                    g.pais = pais;
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} --> CUPS: {g.cups22} El campo 'país' es obligatorio y no está informado.");
                                }
                                //provincia 
                                int celdaProvincia = 42;
                                string valorProvincia = "00";
                                // Validar el campo "provincia"
                                if (workSheet.Cells[f, celdaProvincia].Value != null && Convert.ToString(workSheet.Cells[f, celdaProvincia].Value).Length >= 2)
                                {
                                    valorProvincia = Convert.ToString(workSheet.Cells[f, celdaProvincia].Value).Substring(0, 2).ToUpper();

                                    // Verificar si el valor de la provincia existe en el diccionario "cnmc_p_provincias"
                                    if (cnmc.dic_provincias.ContainsValue(valorProvincia))
                                    {
                                        g.provincia = valorProvincia; // Asignar el valor de la provincia
                                    }
                                    else
                                    {
                                        lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " provincia: el campo provincia no contiene un valor válido.");
                                        //lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + "provincia: el campo provincia es nulo o no tiene una longitud válida (2 caracteres).");
                                    //lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }

                                //municipio 
                                int celdaMunicipio = 43;
                                if (workSheet.Cells[f, celdaMunicipio].Value != null && Convert.ToString(workSheet.Cells[f, celdaMunicipio].Value).Length >= 2)
                                {
                                    //if (cnmc.dic_municipios.ContainsKey(valorProvincia+"|"+Convert.ToString(workSheet.Cells[f, celdaMunicipio].Value).ToLower()))
                                    string cod_municipio;
                                    if (cnmc.dic_municipios.TryGetValue(valorProvincia + "|" + Convert.ToString(workSheet.Cells[f, celdaMunicipio].Value).ToLower(), out cod_municipio))
                                    {
                                        g.municipio = cod_municipio;
                                    }
                                    else
                                    {
                                        lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " municipio: el campo tipo municipios no contiene un valor válido.");
                                        //lista_log.Add(System.Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + " municipio: el campo municipio es nulo o no tiene una longitud válida X(2).");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }

                                //Codigo postal  - 
                                int celdaCP = 44;  // Índice de la columna para el Código Postal
                                if (workSheet.Cells[f, celdaCP].Value != null)
                                {
                                    string valorCP = Convert.ToString(workSheet.Cells[f, celdaCP].Value).Trim();

                                    if (int.TryParse(valorCP, out int codigoPostal))
                                    {
                                        string codigoPostalStr = codigoPostal.ToString().PadLeft(5, '0');  // Completar con ceros a la izquierda si tiene menos de 5 dígitos

                                        if (codigoPostalStr.Length == 5)
                                        {
                                            g.codigo_postal = codigoPostalStr;
                                        }
                                        else
                                        {
                                            lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Código Postal: debe contener exactamente 5 dígitos numéricos.");
                                            //lista_log.Add(Environment.NewLine);
                                            registro_valido = false;
                                        }
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Código Postal: el valor ingresado no es numérico.");
                                        //lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Código Postal: el campo está vacío.");
                                    //lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                                //tipo de via 
                                int celdatipovia = 45;       // Celda para el campo "tipo de via "
                                                             // Validar el campo "tipo de via "
                                if (workSheet.Cells[f, celdatipovia].Value != null && Convert.ToString(workSheet.Cells[f, celdatipovia].Value).Length >= 2)
                                {
                                    string valortipo_via = Convert.ToString(workSheet.Cells[f, celdatipovia].Value).Substring(0, 2).ToUpper();
                                    // Verificar si el valor de la tipo via  en el diccionario "cnmc_p_tipo_via"
                                    if (cnmc.dic_via.ContainsValue(valortipo_via))
                                    {
                                        g.tipo_via = valortipo_via; // Asignar el valor de la tipo via 
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> tipo via : el campo tipo via  no contiene un valor válido.");
                                        lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> tipo via : el campo tipo via es nulo o no tiene una longitud válida (2 caracteres).");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }

                                // nombre via 
                                var valorCelda = workSheet.Cells[f, 46].Value;
                                if (valorCelda != null)
                                {
                                    string nombreVia = valorCelda.ToString().Trim(); // Elimina espacios innecesarios

                                    if (nombreVia.Length > 30)
                                    {
                                        lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " +
                                                      "Nombre de Vía: el texto excede los 30 caracteres y será truncado.");
                                        //lista_log.Add(System.Environment.NewLine);

                                        g.nombre_via = nombreVia.Substring(0, 30); // Corta a 30 caracteres
                                    }
                                    else
                                    {
                                        g.nombre_via = nombreVia;
                                    }
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " +
                                                  "Nombre de Vía: el campo está vacío o es nulo.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                                // numero 
                                int celdaNumero = 47;  // Índice de la columna para "número"
                                if (workSheet.Cells[f, celdaNumero].Value != null)
                                {
                                    string numero = Convert.ToString(workSheet.Cells[f, celdaNumero].Value).Trim();  // Elimina espacios adicionales

                                    if (numero.Length > 5)
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Número: el texto excede los 5 caracteres y será truncado.");
                                        lista_log.Add(Environment.NewLine);

                                        g.numero = numero.Substring(0, 5);  // Trunca el valor a 5 caracteres
                                    }
                                    else if (numero.Length > 0)
                                    {
                                        g.numero = numero;  // Asigna el valor tal cual si tiene entre 1 y 5 caracteres
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Número: el campo está vacío.");
                                        lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Número: el campo está nulo.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }

                            }
                            //  fin direccion - 
                            // entorno  48
                            if (Convert.ToString(workSheet.Cells[f, 48].Value) != "")
                                g.entorno = Convert.ToString(workSheet.Cells[f, 48].Value);

                            // tipo cliente  - 49
                            if (Convert.ToString(workSheet.Cells[f, 49].Value) != "")
                                g.tipo_cliente = Convert.ToString(workSheet.Cells[f, 49].Value);

                            // g.observaciones = Convert.ToString(workSheet.Cells[f, 52].Value);//

                            // TipoDocAportado DireccionUrl
                            c = 49;

                            try
                            {
                                for (int j = 1; j <= 1; j++)  // Bucle para verificar hasta 3 pares / solo 1 par
                                {
                                    c++;
                                    Documentacion doc = new Documentacion();

                                    string tipoDoc = Convert.ToString(workSheet.Cells[f, c].Value)?.Trim();
                                    string urlDireccion = Convert.ToString(workSheet.Cells[f, c + 1].Value)?.Trim();

                                    // Si ambos valores son nulos o vacíos, continuar con el siguiente par
                                    if (string.IsNullOrEmpty(tipoDoc) && string.IsNullOrEmpty(urlDireccion))
                                    {
                                        c++;  // Avanzar dos columnas (porque estamos verificando pares TipoDoc, URL)
                                        continue;
                                    }

                                    // Obtener solo las 2 primeras posiciones de tipoDoc
                                    tipoDoc = !string.IsNullOrEmpty(tipoDoc) ? tipoDoc.Substring(0, Math.Min(2, tipoDoc.Length)) : "00";

                                    // Asignar valores al objeto Documentacion
                                    doc.tipo_doc_aportado = tipoDoc;
                                    doc.direccion_url = urlDireccion;

                                    // Agregar el objeto Documentacion a la lista
                                    g.lista_documentacion.Add(doc);

                                    c++;  // Avanzar la columna de la URL para la siguiente iteración
                                }
                            }
                            catch (Exception ex)
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {f} Columna: {c} --> Error inesperado: {ex.Message}");
                            }

                            ///  ojo  ver-error ---  o.registros- inicio uno excepción de tipo “System.NullReferenceException” 
                            if (registro_valido)
                            {
                                lista_datos_excel_extrasistemas.Add(g);
                                                                
                                if (dic_total_registros.TryGetValue(g.hoja, out o))
                                {
                                    o.registros = o.registros + 1;
                                }
                            }

                      }
                    }  //xls
                    // M101 
                    else if (workSheet.Name == "MOD")  //  xls
                    {

                        //NO PROCESAMOS ESTA PESTAÑA HASTA FINALIZAR IMPLEMENTACION
                        //FORZAMOS BREAK
                        //lista_log.Add("La pestaña MOD no se procesará, está deshabilitada temporalmente");
                        //continue;
                        //
                        f = 2; // Porque la primera fila es la cabecera
                        for (int i = 1; i < 1000000; i++)
                        {
                            registro_valido = true;
                            f++;
                            if (workSheet.Cells[f, 3].Value == null)
                                break;

                            mod_total++;

                            EndesaEntity.extrasistemas.Global g = new EndesaEntity.extrasistemas.Global();
                            g.hoja = workSheet.Name;
                            //
                            #region  Cabecera
                            // empresa 
                            int celdaempresa = 1;
                            if (workSheet.Cells[f, celdaempresa].Value != null && Convert.ToString(workSheet.Cells[f, celdaempresa].Value).Length >= 4)
                            {
                                string valorCeldaE = Convert.ToString(workSheet.Cells[f, celdaempresa].Value);
                                string posibleCodigoEmpresa = valorCeldaE.Substring(0, 4);

                                if (int.TryParse(posibleCodigoEmpresa, out _))
                                {
                                    g.empresa_emisora = posibleCodigoEmpresa;
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Empresa_emisora: el campo código empresa emisora no tiene un formato válido.");
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Empresa_emisora: el campo código empresa emisora es nulo o no tiene una longitud válida X(4).");
                                registro_valido = false;
                            }
                            //
                            //  Distribuidora
                            int celdaDistribuidora = 2;    // Validar Distribuidora (4 caracteres)
                            if (workSheet.Cells[f, celdaDistribuidora].Value != null && Convert.ToString(workSheet.Cells[f, celdaDistribuidora].Value).Length >= 4)
                            {
                                if (int.TryParse(Convert.ToString(workSheet.Cells[f, celdaDistribuidora].Value).Substring(0, 4), out _))
                                {
                                    g.distribuidora = Convert.ToString(workSheet.Cells[f, celdaDistribuidora].Value).Substring(0, 4);
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Distribuidora: el campo código empresa destino (distribuidora) no tiene un formato válido.");
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Distribuidora: el campo código empresa destino (distribuidora) es nulo o no tiene una longitud válida X(4).");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //
                            //  cups22
                            int celdaCUPS = 3;    // Validar CUPS (22 caracteres) 
                            if (workSheet.Cells[f, celdaCUPS].Value != null && Convert.ToString(workSheet.Cells[f, celdaCUPS].Value).Length >= 22)
                            {
                                g.cups22 = Convert.ToString(workSheet.Cells[f, celdaCUPS].Value).Substring(0, 22);
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> CUPS: el campo CUPS es nulo o no tiene una longitud válida X(22).");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }

                            #endregion
                            //  Modificacion de ATR
                            #region DatosSolicitud
                            // tipo modificacion 
                            int celdatipomodificacion = 4;
                            if (workSheet.Cells[f, celdatipomodificacion].Value != null && Convert.ToString(workSheet.Cells[f, celdatipomodificacion].Value).Length >= 1)
                            {
                                g.tipo_modificacion = Convert.ToString(workSheet.Cells[f, celdatipomodificacion].Value).Substring(0, 1);
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Tipo_modificacion: el campo tipo de modificación es nulo o no tiene una longitud válida X(1).");
                                registro_valido = false;
                            }
                            //  tipo solicitud administrativa 
                            int celdatiposolicitudadministrativa = 5;
                            if (g.tipo_modificacion == "A" || g.tipo_modificacion == "S")
                            {
                                if (workSheet.Cells[f, celdatiposolicitudadministrativa].Value != null && Convert.ToString(workSheet.Cells[f, celdatiposolicitudadministrativa].Value).Length >= 1)
                                {
                                    g.tipo_solicitud_administrativa = Convert.ToString(workSheet.Cells[f, celdatiposolicitudadministrativa].Value).Substring(0, 1);
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Tipo_solicitud_administrativa: el campo es requerido para modificaciones de tipo 'A' o 'S' y es nulo o tiene una longitud inválida X(1).");
                                    registro_valido = false;
                                }
                            }
                            //

                            int celdaCNAE = 6;   // Variable para la celda de CNAE
                            if (workSheet.Cells[f, celdaCNAE].Value != null && Convert.ToString(workSheet.Cells[f, celdaCNAE].Value).Length >= 4)
                            {
                                g.cnae = Convert.ToString(workSheet.Cells[f, celdaCNAE].Value).Substring(0, 4);
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> CNAE: el campo CNAE es nulo o no tiene una longitud válida X(4).");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }

                            // IndEsencial
                            int celdaEsen = 7;
                            if (workSheet.Cells[f, celdaEsen].Value != null && Convert.ToString(workSheet.Cells[f, celdaEsen].Value).Length >= 2)
                            {
                                g.indesencial = Convert.ToString(workSheet.Cells[f, celdaEsen].Value).Substring(0, 2);
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> IndEsencial: el campo Esencial es nulo o no tiene una longitud válida X(4).");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }

                            //IndActivacion ---------------------------------------------------
                            //int celdaIndActivacion = 8;   

                            //if (workSheet.Cells[f, celdaIndActivacion].Value != null && Convert.ToString(workSheet.Cells[f, celdaIndActivacion].Value).Length >= 1)
                            //{
                            //    string valorIndActivacion = Convert.ToString(workSheet.Cells[f, celdaIndActivacion].Value).Substring(0, 1).ToUpper();

                            //    if (cnmc.dic_indicativo_activacion.ContainsValue(valorIndActivacion))
                            //    {
                            //        g.ind_activacion = valorIndActivacion;
                            //    }
                            //    else
                            //    {
                            //        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> IndActivacion: el campo IndActivacion no contiene un valor válido.");
                            //        // lista_log.Add(Environment.NewLine);
                            //        registro_valido = false;
                            //    }
                            //}
                            //else
                            //{
                            //    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> IndActivacion: el campo IndActivacion es nulo o no tiene una longitud válida X(1).");
                            //    //lista_log.Add(Environment.NewLine);
                            //    registro_valido = false;
                            //}
                            // ------------------------------------------------------------------------
                            int celdaIndActivacion = 8;
                            object celdaValor = workSheet.Cells[f, celdaIndActivacion].Value;

                            string tipoModificacion = g.tipo_modificacion;
                            string tipoSolicitud = g.tipo_solicitud_administrativa;

                            string valorIndActivacion = celdaValor != null
                                ? Convert.ToString(celdaValor).Trim().ToUpper().Substring(0, 1)
                                : null;

                            // Validar si el valor está presente y es válido
                            if (!string.IsNullOrEmpty(valorIndActivacion) && cnmc.dic_indicativo_activacion.ContainsValue(valorIndActivacion))
                            {
                                g.ind_activacion = valorIndActivacion;
                            }
                            else
                            {
                                // Asignar valor por defecto si no está informado o no es válido
                                if (tipoModificacion == "S")
                                {
                                    if (tipoSolicitud == "A" || tipoSolicitud == "S" || tipoSolicitud == "C")
                                    {
                                        g.ind_activacion = "A";
                                    }
                                    else if (tipoSolicitud == "P")
                                    {
                                        g.ind_activacion = "L";
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> IndActivacion: tipoSolicitudAdministrativa no reconocido para tipoModificacion=S.");
                                        registro_valido = false;
                                    }
                                }
                                else if (tipoModificacion == "M" || tipoModificacion == "B")
                                {
                                    g.ind_activacion = "A";
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> IndActivacion: tipoModificacion no reconocido. No se puede asignar valor por defecto.");
                                    registro_valido = false;
                                }

                                // Verificar si el valor por defecto es válido según el diccionario
                                if (!cnmc.dic_indicativo_activacion.ContainsValue(g.ind_activacion))
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> IndActivacion: el valor asignado por defecto '{g.ind_activacion}' no es válido según el diccionario.");
                                    registro_valido = false;
                                }
                            }

                            //
                            int celdaFechaActivacion = 9;  // Variable para la celda de fecha_activacion
                            if (g.ind_activacion == "F")
                            {
                                var valorFecha = workSheet.Cells[f, celdaFechaActivacion].Value;

                                if (valorFecha != null)
                                {
                                    string fechaStr = Convert.ToString(valorFecha).Trim();

                                    if (DateTime.TryParse(fechaStr, out DateTime fechaParsed))
                                    {
                                        // Verifica que el formato sea exactamente AAAA-MM-DD
                                        if (fechaParsed.ToString("yyyy-MM-dd") == fechaStr)
                                        {
                                            g.fecha_activacion = fechaParsed;
                                        }
                                        else
                                        {
                                            lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> FechaPrevistaAccion: el campo fecha_activación no tiene el formato AAAA-MM-DD.");
                                            lista_log.Add(Environment.NewLine);
                                            registro_valido = false;
                                        }
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> FechaPrevistaAccion: el campo fecha_activación no es una fecha válida.");
                                        lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> FechaPrevistaAccion: el campo fecha_activación está vacío.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            //  solicitud modificacion tension
                            int celdasolicitudmodtension = 10;

                            if (workSheet.Cells[f, celdasolicitudmodtension].Value != null && Convert.ToString(workSheet.Cells[f, celdasolicitudmodtension].Value).Length >= 1)
                            {
                                string valorTension = Convert.ToString(workSheet.Cells[f, celdasolicitudmodtension].Value).Substring(0, 1).ToUpper();

                                if (valorTension == "N" || valorTension == "S")
                                {
                                    g.solicitud_modificacion_tension = valorTension;
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Solicitud_modificacion_tension: el valor debe ser 'N' o 'S'.");
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Solicitud_modificacion_tension: el campo es nulo o no tiene una longitud válida X(1).");
                                registro_valido = false;
                            }

                            //g.solicitud_modificacion_tension = Convert.ToString(workSheet.Cells[f, 12].Value).Substring(0, 1);
                            //if (g.solicitud_modificacion_tension == "S")
                            //    g.tension_solicitada = Convert.ToString(workSheet.Cells[f, 13].Value).Substring(0, 2);

                            // Tension solicitada --      tabla  -   mysql cnmc_p_tensiones
                            int celdatensionsolicitada = 11;

                            if (g.solicitud_modificacion_tension == "S")
                            {
                                if (workSheet.Cells[f, celdatensionsolicitada].Value != null && Convert.ToString(workSheet.Cells[f, celdatensionsolicitada].Value).Length >= 2)
                                {
                                    string valorTension = Convert.ToString(workSheet.Cells[f, celdatensionsolicitada].Value).Substring(0, 2);

                                    if (cnmc.dic_tensiones.ContainsValue(valorTension))
                                    {
                                        g.tension_solicitada = valorTension;
                                    }
                                    else
                                    {
                                        lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Tension_solicitada: el valor '" + valorTension + "' no se encuentra en el diccionario de tensiones válido.");
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> Tension_solicitada: campo requerido porque solicitud_modificacion_tension es 'S', pero es nulo o no tiene una longitud válida X(2).");
                                    registro_valido = false;
                                }
                            }

                            #endregion
                            //  
                            #region contrato
                            //tipo de contrato 12
                            // dic_contrato_atr = Carga_Tabla_CNMC("cnmc_p_tipo_contrato_atr");

                            // Declarar variable para la celda del contrato ATR
                            var celdaTipoContratoATR = workSheet.Cells[f, 12].Value;

                            if (celdaTipoContratoATR != null && Convert.ToString(celdaTipoContratoATR).Length >= 2)
                            {
                                string tipoContratoa = Convert.ToString(celdaTipoContratoATR).Substring(0, 2).ToUpper();

                                // Ojo: la clave en el diccionario está al revés en la descripción
                                if (cnmc.dic_contrato_atr.ContainsValue(tipoContratoa))
                                {
                                    g.tipo_contrato_atr = tipoContratoa;
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> tipo_contrato: el campo tipo_contrato_at no contiene un valor válido.");
                                    // lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> Fila: " + i + " tipo_contrato : el campo tipo_contrato_at es nulo o no tiene una longitud válida (X2).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }

                            // fecha de finalizacion  13
                            var celdaFechaFinalizacion = workSheet.Cells[f, 13].Value;
                            string tipoContrato = g.tipo_contrato_atr;

                            // Validar si el tipo de contrato requiere fecha de finalización
                            if (tipoContrato == "02" || tipoContrato == "03" || tipoContrato == "09" || tipoContrato == "10")
                            {
                                if (celdaFechaFinalizacion == null || string.IsNullOrWhiteSpace(Convert.ToString(celdaFechaFinalizacion)))
                                {
                                    registro_valido = false;
                                    lista_log.Add("Hoja: " + g.hoja + " --> CUPS: " + g.cups22 +
                                        " --> El campo fecha finalización es obligatorio para contratos tipo '02', '03', '09' o '10'.");
                                }
                                else
                                {
                                    // Validar formato de fecha antes de asignar
                                    if (DateTime.TryParse(Convert.ToString(celdaFechaFinalizacion), out DateTime fechaFin))
                                    {
                                        g.fecha_finalizacion = fechaFin;
                                    }
                                    else
                                    {
                                        registro_valido = false;
                                        lista_log.Add("Hoja: " + g.hoja + " --> CUPS: " + g.cups22 +
                                            " --> El campo fecha finalización no tiene un formato de fecha válido.");
                                    }
                                }
                            }
                            else
                            {
                                // Si no es uno de los tipos requeridos, no se debe informar fecha de finalización
                                if (celdaFechaFinalizacion != null && !string.IsNullOrWhiteSpace(Convert.ToString(celdaFechaFinalizacion)))
                                {
                                    registro_valido = false;
                                    lista_log.Add("Hoja: " + g.hoja + " --> CUPS: " + g.cups22 +
                                        " --> El campo fecha finalización no debe estar informado para este tipo de contrato (" + tipoContrato + ").");
                                }
                            }
                            // TipoCUPS 14
                            int celdatipocupsM = 14;
                            var valorCeldatc = workSheet.Cells[f, celdatipocupsM].Value;

                            if (valorCeldatc != null && Convert.ToString(valorCeldatc).Length >= 2)
                            {
                                // Se obtiene el valor de las primeras dos letras en mayúsculas
                                string celdatipocups = Convert.ToString(valorCeldatc).Substring(0, 2).ToUpper();
                                // Asignar el valor directamente
                                g.tipocups = celdatipocups;
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} tipocups: el campo tipocups es nulo o no tiene una longitud válida.");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }

                            //cau 15
                            int celdaCauM = 15;
                            var valorCeldaCau = workSheet.Cells[f, celdaCauM].Value;

                            if (valorCeldaCau != null && Convert.ToString(valorCeldaCau).Length == 26)
                            {
                                // Asignar el valor directamente
                                g.cau = Convert.ToString(valorCeldaCau);
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} cau: el campo cau es nulo o no tiene una longitud válida.");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //

                            // tipo autoconsumo  16
                            var celdaTipoAutoconsumo = workSheet.Cells[f, 16].Value;
                            if (celdaTipoAutoconsumo != null && Convert.ToString(celdaTipoAutoconsumo).Length >= 2)
                            {
                                string tipoAutoconsumoStr = Convert.ToString(celdaTipoAutoconsumo).Substring(0, 2).ToUpper();
                                // Validación por descripción en el diccionario (ojo con ContainsValue)
                                if (cnmc.dic_autoconsumo.ContainsValue(tipoAutoconsumoStr))
                                {
                                    g.tipo_autoconsumo = tipoAutoconsumoStr;
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> tipo_autoconsumo: el campo tipo_autoconsumo no contiene un valor válido.");
                                    // lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> Fila: " + i + " tipo_autoconsumo: el campo tipo_autoconsumo es nulo o no tiene una longitud válida (X2).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }


                            // TipoSubseccion  
                            int celdaTiposubseccionM = 17;
                            var valorCeldaTiposubseccion = workSheet.Cells[f, celdaTiposubseccionM].Value;

                            if (valorCeldaTiposubseccion != null && Convert.ToString(valorCeldaTiposubseccion).Length >= 2)
                            {
                                // Se obtiene el valor de las primeras dos letras en mayúsculas
                                string tiposubseccionValor = Convert.ToString(valorCeldaTiposubseccion).Substring(0, 2).ToUpper();
                                // Asignar el valor directamente
                                g.tipo_subseccion = tiposubseccionValor;
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} tiposubseccion: el campo tiposubseccion es nulo o no tiene una longitud válida.");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //PotInstaladaGen   18
                            int celdaPotInstaladaGen = 18;
                            var valorCeldaPotInstaladaGen = workSheet.Cells[f, celdaPotInstaladaGen].Value;

                            if (valorCeldaPotInstaladaGen != null && long.TryParse(Convert.ToString(valorCeldaPotInstaladaGen), out long potInstaladaGenValor) && Convert.ToString(valorCeldaPotInstaladaGen).Length <= 14)
                            {
                                // Asignar el valor directamente
                                g.potinstaladagen = (int)potInstaladaGenValor;
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} --> Fila: {i} potInstaladaGen: el campo potInstaladaGen es nulo, no es numérico o excede las 14 posiciones.");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //
                            #endregion

                            #region potencia
                            // Tarifa -  19
                            var celdaTarifa = workSheet.Cells[f, 19].Value;
                            if (celdaTarifa != null && Convert.ToString(celdaTarifa).Length >= 3)
                            {
                                string tarifaStr = Convert.ToString(celdaTarifa).Substring(0, 3).ToUpper();
                                // Ojo: la clave es la descripción (ContainsValue)
                                if (cnmc.dic_tarifa_atr.ContainsValue(tarifaStr))
                                {
                                    g.tarifa = tarifaStr;
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> tarifa atr: el campo tarifa atr no contiene un valor válido.");
                                    lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> Fila: " + i + " tarifa atr: el campo tarifa atr es nulo o no tiene una longitud válida (X3).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }

                            // Validación de las potencias contratadas en vatios (máximo 6)
                            int potenciasInformadas = 0;
                            bool esTarifa20TD = (g.tarifa == "2.0TD");

                            for (int pot = 0; pot < 6; pot++)
                            {
                                var celdaPotencia = workSheet.Cells[f, 20 + pot].Value;
                                string potenciaStr = Convert.ToString(celdaPotencia);

                                if (!string.IsNullOrWhiteSpace(potenciaStr) && int.TryParse(potenciaStr, out int potencia))
                                {
                                    if (potencia >= 0)
                                    {
                                        g.potencias[pot] = potencia;
                                        potenciasInformadas++;
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Potencia {pot + 1}: el valor debe ser un número positivo.");
                                        //lista_log.Add(Environment.NewLine);  CUPS: {g.cups22}
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Potencia {pot + 1}: el valor no es un número válido o está en blanco.");
                                    //lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }

                            // Validación mínima: al menos 3 potencias deben estar informadas
                            if (potenciasInformadas < 3)
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Deben estar informadas al menos 3 potencias.");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }

                            // Validación de orden creciente de potencias (solo si no es tarifa 2.0TD)
                            if (!esTarifa20TD)
                            {
                                for (int pot = 1; pot < 6; pot++)
                                {
                                    // Si alguna potencia no está informada, se salta la validación de orden
                                    if (g.potencias[pot] == 0 || g.potencias[pot - 1] == 0) continue;

                                    if (g.potencias[pot] < g.potencias[pot - 1])
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Las potencias deben ser crecientes. Potencia {pot + 1} ({g.potencias[pot]}) es menor que Potencia {pot} ({g.potencias[pot - 1]}).");
                                        lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                        break;
                                    }
                                }
                            }

                            // Validación del campo modo_control_potencia - 
                            var celdaModoControl = workSheet.Cells[f, 26].Value;

                            if (celdaModoControl != null && Convert.ToString(celdaModoControl).Length >= 1)
                            {
                                string modoControlStr = Convert.ToString(celdaModoControl).Substring(0, 1).ToUpper();

                                g.modo_control_potencia = modoControlStr;

                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> El campo modo_control_potencia está vacío o no tiene una longitud válida.");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }

                            #endregion

                            #region Cliente
                            // Validación del campo tipo_identificador - 
                            var celdaIdentificador = workSheet.Cells[f, 27].Value;
                            if (celdaIdentificador != null)
                            {
                                // Normalizar el valor de la celda
                                string descripcion = Convert.ToString(celdaIdentificador)?.Trim().ToUpperInvariant();

                                if (!string.IsNullOrEmpty(descripcion))
                                {
                                    // Normalizar claves del diccionario eliminando espacios y asegurando mayúsculas
                                    var dic_normalizado = cnmc.dic_identificador
                                        .ToDictionary(kv => kv.Key.Trim().ToUpperInvariant(), kv => kv.Value);

                                    // Buscar coincidencia directa
                                    if (!dic_normalizado.TryGetValue(descripcion, out string valorEncontrado))
                                    {
                                        // Intento adicional: Buscar con coincidencias parciales si las claves son listas
                                        foreach (var kvp in dic_normalizado)
                                        {
                                            var clavesDiccionario = kvp.Key.Split(new[] { ',', '-', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                                           .Select(s => s.Trim().ToUpperInvariant());

                                            var clavesEntrada = descripcion.Split(new[] { ',', '-', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                                           .Select(s => s.Trim().ToUpperInvariant());

                                            if (clavesDiccionario.Intersect(clavesEntrada).Any()) // Coincidencia parcial
                                            {
                                                valorEncontrado = kvp.Value;
                                                break;
                                            }
                                        }
                                    }

                                    // Verificar si encontramos el valor
                                    if (!string.IsNullOrEmpty(valorEncontrado))
                                    {
                                        if (valorEncontrado.Length >= 2)
                                        {
                                            g.tipo_identificador = valorEncontrado.Substring(valorEncontrado.Length - 2).ToUpper();
                                        }
                                        else
                                        {
                                            lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El valor asociado a '{descripcion}' en el diccionario es inválido ({valorEncontrado}).");
                                            lista_log.Add(Environment.NewLine);
                                            registro_valido = false;
                                        }
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: La descripción '{descripcion}' no se encontró en el diccionario.");
                                        lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El valor en la celda es inválido (vacío o solo espacios).");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El campo es nulo.");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //
                            // Validación campo n_identificador - 28
                            var celdaIdentifica = workSheet.Cells[f, 28].Value;

                            if (celdaIdentifica != null && Convert.ToString(celdaIdentifica).Trim().Length >= 1)
                            {
                                string valorIdentificador = Convert.ToString(celdaIdentifica).Trim();

                                // Validación de longitud máxima de 14 caracteres
                                if (valorIdentificador.Length <= 14)
                                {
                                    g.n_identificador = valorIdentificador.ToUpper();
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> n_identificador: El valor excede la longitud máxima permitida (14 caracteres).");
                                    //lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> n_identificador: El campo es nulo o está vacío.");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }

                            // tipo  de persona 
                            var tipoPersona = Convert.ToString(workSheet.Cells[f, 29].Value);
                            // Comprobar si el tipo de persona no está vacío
                            if (!string.IsNullOrEmpty(tipoPersona))
                            {
                                g.tipo_persona = tipoPersona.Substring(0, 1);

                                if (g.tipo_persona == "J")
                                {
                                    var razonSocial = Convert.ToString(workSheet.Cells[f, 30].Value);
                                    g.razon_social = razonSocial?.Substring(0, Math.Min(razonSocial.Length, 50));
                                }
                                else
                                {
                                    var nombre = Convert.ToString(workSheet.Cells[f, 31].Value);
                                    var apellido = Convert.ToString(workSheet.Cells[f, 32].Value);

                                    g.nombre_de_pila = nombre?.Substring(0, Math.Min(nombre.Length, 50));
                                    g.primer_apellido = apellido?.Substring(0, Math.Min(apellido.Length, 50));
                                }
                            }

                            // telefono 33
                            var telefonoRaw = Convert.ToString(workSheet.Cells[f, 33].Value);
                            // Validar que el teléfono no esté vacío y tenga entre 6 y 12 dígitos numéricos
                            if (!string.IsNullOrEmpty(telefonoRaw))
                            {
                                // Eliminar espacios y caracteres no numéricos si se desea (opcional)
                                var telefono = new string(telefonoRaw.Where(char.IsDigit).ToArray());

                                if (telefono.Length >= 6 && telefono.Length <= 12)
                                {
                                    g.telefono = telefono;
                                    g.tlf_contracto = telefono;
                                }
                                else
                                {
                                    registro_valido = false;
                                    lista_log.Add($"Hoja: {g.hoja} --> CUPS: {g.cups22} El campo teléfono debe tener entre 6 y 12 dígitos.");
                                }
                            }
                            else
                            {
                                registro_valido = false;
                                lista_log.Add($"Hoja: {g.hoja} --> CUPS: {g.cups22} El campo teléfono no tiene un valor válido.");
                            }

                            // persona de contacto 34

                            var personaContactoRaw = Convert.ToString(workSheet.Cells[f, 34].Value);

                            if (!string.IsNullOrWhiteSpace(personaContactoRaw))
                            {
                                var personaContacto = personaContactoRaw.Substring(0, Math.Min(personaContactoRaw.Length, 150));
                                g.persona_contacto = personaContacto;
                                g.contacto = personaContacto;
                            }
                            else
                            {
                                registro_valido = false;
                                lista_log.Add($"Hoja: {g.hoja} --> CUPS: {g.cups22} El campo 'persona de contacto' es obligatorio y está vacío.");
                            }
                            #endregion

                            // indicador tipo direccion 
                            var indicadorRaw = Convert.ToString(workSheet.Cells[f, 35].Value);
                            if (!string.IsNullOrWhiteSpace(indicadorRaw))
                            {
                                g.indicador_tipo_direccion = indicadorRaw.Substring(0, 1);
                            }
                            else
                            {
                                registro_valido = false;
                                lista_log.Add($"Hoja: {g.hoja} --> CUPS: {g.cups22} El campo 'IndicadorTipoDireccion' no tiene un valor válido.");
                            }

                            #region Direccion
                            //pais
                            string pais = Convert.ToString(workSheet.Cells[f, 36].Value)?.Trim();   // Obtener valor de la celda y convertirlo a string eliminando espacios
                                                                                                    // Validar y asignar o registrar en lista de errores
                            if (!string.IsNullOrEmpty(pais))
                            {
                                g.pais = pais;
                            }
                            else
                            {
                                registro_valido = false;
                                lista_log.Add($"Hoja: {g.hoja} --> CUPS: {g.cups22} El campo 'país' es obligatorio y no está informado.");
                            }


                            //privincia
                            int celdaProvincia = 37;
                            string valorProvincia = "00";
                            // Validar el campo "provincia"
                            if (workSheet.Cells[f, celdaProvincia].Value != null && Convert.ToString(workSheet.Cells[f, celdaProvincia].Value).Length >= 2)
                            {
                                valorProvincia = Convert.ToString(workSheet.Cells[f, celdaProvincia].Value).Substring(0, 2).ToUpper();

                                // Verificar si el valor de la provincia existe en el diccionario "cnmc_p_provincias"
                                if (cnmc.dic_provincias.ContainsValue(valorProvincia))
                                {
                                    g.provincia = valorProvincia; // Asignar el valor de la provincia
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " provincia: el campo provincia no contiene un valor válido.");
                                    //lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + "provincia: el campo provincia es nulo o no tiene una longitud válida (2 caracteres).");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }

                            //municipio
                            int celdaMunicipio = 38;
                            if (workSheet.Cells[f, celdaMunicipio].Value != null && Convert.ToString(workSheet.Cells[f, celdaMunicipio].Value).Length >= 2)
                            {
                                //if (cnmc.dic_municipios.ContainsKey(valorProvincia+"|"+Convert.ToString(workSheet.Cells[f, celdaMunicipio].Value).ToLower()))
                                string cod_municipio;
                                if (cnmc.dic_municipios.TryGetValue(valorProvincia + "|" + Convert.ToString(workSheet.Cells[f, celdaMunicipio].Value).ToLower(), out cod_municipio))
                                {
                                    g.municipio = cod_municipio;
                                }
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " municipio: el campo tipo municipios no contiene un valor válido.");
                                    lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " --> " + " Fila: " + i + " municipio: el campo municipio es nulo o no tiene una longitud válida X(2).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }


                            // codigo postal
                            int celdaCP = 39;  // Índice de la columna para el Código Postal

                            if (workSheet.Cells[f, celdaCP].Value != null)
                            {
                                string valorCP = Convert.ToString(workSheet.Cells[f, celdaCP].Value).Trim();

                                if (int.TryParse(valorCP, out int codigoPostal))
                                {
                                    string codigoPostalStr = codigoPostal.ToString().PadLeft(5, '0');  // Completar con ceros a la izquierda si tiene menos de 5 dígitos

                                    if (codigoPostalStr.Length == 5)
                                    {
                                        g.codigo_postal = codigoPostalStr;
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Código Postal: debe contener exactamente 5 dígitos numéricos.");
                                        lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Código Postal: el valor ingresado no es numérico.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Código Postal: el campo está vacío.");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //
                            //tipo via 
                            int celdatipovia = 40;
                            // Validar el campo "tipo de via "
                            if (workSheet.Cells[f, celdatipovia].Value != null && Convert.ToString(workSheet.Cells[f, celdatipovia].Value).Length >= 2)
                            {
                                string valortipo_via = Convert.ToString(workSheet.Cells[f, celdatipovia].Value).Substring(0, 2).ToUpper();

                                // Verificar si el valor de la tipo via  en el diccionario "cnmc_p_tipo_via"
                                if (cnmc.dic_via.ContainsValue(valortipo_via))
                                {
                                    g.tipo_via = valortipo_via; // Asignar el valor de la tipo via 
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> tipo via : el campo tipo via  no contiene un valor válido.");
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> tipo via : el campo tipo via es nulo o no tiene una longitud válida (2 caracteres).");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //Nombre de via  
                            var valorCelda = workSheet.Cells[f, 41].Value;
                            if (valorCelda != null)
                            {
                                string nombreVia = valorCelda.ToString().Trim(); // Elimina espacios innecesarios

                                if (nombreVia.Length > 30)
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " +
                                                  "Nombre de Vía: el texto excede los 30 caracteres y será truncado.");
                                    // lista_log.Add(System.Environment.NewLine);

                                    g.nombre_via = nombreVia.Substring(0, 30); // Corta a 30 caracteres
                                }
                                else
                                {
                                    g.nombre_via = nombreVia;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " +
                                              "Nombre de Vía: el campo está vacío o es nulo.");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //numero 
                            int celdaNumero = 42;  // Índice de la columna para "número"
                            if (workSheet.Cells[f, celdaNumero].Value != null)
                            {
                                string numero = Convert.ToString(workSheet.Cells[f, celdaNumero].Value).Trim();  // Elimina espacios adicionales

                                if (numero.Length > 5)
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Número: el texto excede los 5 caracteres y será truncado.");
                                    lista_log.Add(Environment.NewLine);

                                    g.numero = numero.Substring(0, 5);  // Trunca el valor a 5 caracteres
                                }
                                else if (numero.Length > 0)
                                {
                                    g.numero = numero;  // Asigna el valor tal cual si tiene entre 1 y 5 caracteres
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Número: el campo está vacío.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Número: el campo está nulo.");
                                lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }

                            #endregion
                            // Entorno 
                            if (Convert.ToString(workSheet.Cells[f, 43].Value) != "")
                                g.entorno = Convert.ToString(workSheet.Cells[f, 43].Value);
                            // Tipo Cliente
                            if (Convert.ToString(workSheet.Cells[f, 44].Value) != "")

                                g.tipo_cliente = Convert.ToString(workSheet.Cells[f, 44].Value);

                            c = 44;

                            // TipoDocAportado DireccionUrl
                            try
                            {
                                for (int j = 1; j <= 1; j++)  // Bucle para verificar hasta 3 pares / solo 1 par
                                {
                                    c++;
                                    Documentacion doc = new Documentacion();

                                    string tipoDoc = Convert.ToString(workSheet.Cells[f, c].Value)?.Trim();
                                    string urlDireccion = Convert.ToString(workSheet.Cells[f, c + 1].Value)?.Trim();

                                    // Si ambos valores son nulos o vacíos, continuar con el siguiente par
                                    if (string.IsNullOrEmpty(tipoDoc) && string.IsNullOrEmpty(urlDireccion))
                                    {
                                        c++;  // Avanzar dos columnas (porque estamos verificando pares TipoDoc, URL)
                                        continue;
                                    }

                                    // Obtener solo las 2 primeras posiciones de tipoDoc
                                    tipoDoc = !string.IsNullOrEmpty(tipoDoc) ? tipoDoc.Substring(0, Math.Min(2, tipoDoc.Length)) : "00";

                                    // Asignar valores al objeto Documentacion
                                    doc.tipo_doc_aportado = tipoDoc;
                                    doc.direccion_url = urlDireccion;

                                    // Agregar el objeto Documentacion a la lista
                                    g.lista_documentacion.Add(doc);

                                    c++;  // Avanzar la columna de la URL para la siguiente iteración
                                }
                            }
                            catch (Exception ex)
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {f} Columna: {c} --> Error inesperado: {ex.Message}");
                            }


                             if (registro_valido)
                             {
                                lista_datos_excel_extrasistemas.Add(g);
                                if (dic_total_registros.TryGetValue(g.hoja, out o))
                                    o.registros = o.registros + 1;
                             }

                        }

                    }   // xls
                    // B101    xls
                    else if (workSheet.Name == "BAJA")  //   xls
                    {
                        //NO PROCESAMOS ESTA PESTAÑA HASTA FINALIZAR IMPLEMENTACION
                        //FORZAMOS BREAK
                       // lista_log.Add("La pestaña BAJA  no se procesará, está deshabilitada temporalmente");
                        //continue;
                        //
                        f = 1; // Porque la primera fila es la cabecera
                        for (int i = 1; i < 1000000; i++)
                        {
                            registro_valido = true;
                            f++;
                            if (workSheet.Cells[f, 3].Value == null)
                                break;

                            baja_total++;
                            #region cabecera
                            EndesaEntity.extrasistemas.Global g = new EndesaEntity.extrasistemas.Global();
                            g.hoja = workSheet.Name;
                            //  CodigoREEEmpresaEmisora - Obligatorio X(4) - STRING NUMERICO  - 
                            if (workSheet.Cells[f, 1].Value != null && Convert.ToString(workSheet.Cells[f, 1].Value).Length >= 4)
                            {
                                if (int.TryParse(Convert.ToString(workSheet.Cells[f, 1].Value).Substring(0, 4), out _))

                                    g.empresa_emisora = Convert.ToString(workSheet.Cells[f, 1].Value).Substring(0, 4);
                                else
                                {
                                    lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " Empresa_emisora: el campo código empresa emisora no tiene un formato válido.");
                                    //lista_log.Add(System.Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add("Hoja: " + g.hoja + " Fila: " + i + " --> " + " Empresa_emisora: el campo código empresa emisora es nulo o no tiene una longitud válida X(4).");
                                //lista_log.Add(System.Environment.NewLine);
                                registro_valido = false;
                            }
                            //  Distribuidora
                            int celdaDistribuidora = 2;  // Variable para la celda de Distribuidora
                            if (workSheet.Cells[f, celdaDistribuidora].Value != null && Convert.ToString(workSheet.Cells[f, celdaDistribuidora].Value).Length >= 4)
                            {
                                if (int.TryParse(Convert.ToString(workSheet.Cells[f, celdaDistribuidora].Value).Substring(0, 4), out _))
                                {
                                    g.distribuidora = Convert.ToString(workSheet.Cells[f, celdaDistribuidora].Value).Substring(0, 4);
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Distribuidora: el campo código empresa destino (distribuidora) no tiene un formato válido.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Distribuidora: el campo código empresa destino (distribuidora) es nulo o no tiene una longitud válida X(4).");
                                // lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //  cups22
                            int celdaCUPS = 3;   // Variable para la celda de CUPS
                            if (workSheet.Cells[f, celdaCUPS].Value != null && Convert.ToString(workSheet.Cells[f, celdaCUPS].Value).Length >= 22)
                            {
                                g.cups22 = Convert.ToString(workSheet.Cells[f, celdaCUPS].Value).Substring(0, 22);
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> CUPS: el campo CUPS es nulo o no tiene una longitud válida X(22).");
                                // lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            

                            #endregion
                            //
                            #region DatosSolicitud
                            //
                            // Variable para la celda de IndActivacion
                            int celdaIndActivacion = 4;
                            if (workSheet.Cells[f, celdaIndActivacion].Value != null && Convert.ToString(workSheet.Cells[f, celdaIndActivacion].Value).Length >= 1)
                            {
                                string indActivacion = Convert.ToString(workSheet.Cells[f, celdaIndActivacion].Value).Substring(0, 1).ToUpper();

                                if (cnmc.dic_indicativo_activacion.ContainsValue(indActivacion))
                                {
                                    g.ind_activacion = indActivacion;
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> IndActivacion: el campo IndActivacion no contiene un valor válido.");
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> IndActivacion: el campo IndActivacion es nulo o no tiene una longitud válida X(1).");
                                registro_valido = false;
                            }

                            if (g.ind_activacion == "F")
                            {
                                // Nos aseguramos que el campo contiene una fecha válida
                                string fechaBj = Convert.ToString(workSheet.Cells[f, 5].Value);

                                if (DateTime.TryParseExact(fechaBj, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out DateTime fechaActivacion))
                                {
                                    g.fecha_activacion = fechaActivacion;
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Fechaactivacion: el campo fecha_activación no tiene un campo fecha válido AAAA-MM-DD.");
                                    registro_valido = false;
                                }
                            }


                            //
                            int celdaMotivoBaja = 6;  // Variable para la celda de Motivo Baja
                            if (workSheet.Cells[f, celdaMotivoBaja].Value != null && Convert.ToString(workSheet.Cells[f, celdaMotivoBaja].Value).Length >= 2)
                            {
                                if (int.TryParse(Convert.ToString(workSheet.Cells[f, celdaMotivoBaja].Value).Substring(0, 2), out _))
                                {
                                    g.motivo_baja = Convert.ToString(workSheet.Cells[f, celdaMotivoBaja].Value).Substring(0, 2);
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Motivo Baja: el campo Motivo Baja no tiene un formato válido.");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Motivo Baja: el campo Motivo Baja es nulo o no tiene una longitud válida X(2).");
                                // lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }
                            //g.motivo_baja = Convert.ToString(workSheet.Cells[f, 6].Value).Substring(0, 2);
                            //
                            #endregion

                            #region Cliente
                            // tipo indentificador  buiscar en diccionario 
                            if (workSheet.Cells[f, 7].Value != null)
                            {
                                // Normalizar el valor de la celda
                                string descripcion = Convert.ToString(workSheet.Cells[f, 7].Value)?.Trim().ToUpperInvariant();

                                if (!string.IsNullOrEmpty(descripcion))
                                {
                                    // Normalizar claves del diccionario eliminando espacios y asegurando mayúsculas
                                    var dic_normalizado = cnmc.dic_identificador
                                        .ToDictionary(kv => kv.Key.Trim().ToUpperInvariant(), kv => kv.Value);

                                    // Buscar la clave en el diccionario
                                    if (!dic_normalizado.TryGetValue(descripcion, out string valorEncontrado))
                                    {
                                        // Intento adicional: Buscar con coincidencias parciales si las claves son listas
                                        foreach (var kvp in dic_normalizado)
                                        {
                                            var clavesDiccionario = kvp.Key.Split(new[] { ',', '-', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                                           .Select(s => s.Trim().ToUpperInvariant());

                                            var clavesEntrada = descripcion.Split(new[] { ',', '-', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                                           .Select(s => s.Trim().ToUpperInvariant());

                                            if (clavesDiccionario.Intersect(clavesEntrada).Any()) // Coincidencia parcial
                                            {
                                                valorEncontrado = kvp.Value;
                                                break;
                                            }
                                        }
                                    }

                                    // Verificar si encontramos el valor  - tipo de indetificador 
                                    if (!string.IsNullOrEmpty(valorEncontrado))
                                    {
                                        if (valorEncontrado.Length >= 2)
                                        {
                                            g.tipo_identificador = valorEncontrado.Substring(valorEncontrado.Length - 2).ToUpper();
                                        }
                                        else
                                        {
                                            lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El valor asociado a '{descripcion}' en el diccionario es inválido ({valorEncontrado}).");
                                            //lista_log.Add(Environment.NewLine);
                                            registro_valido = false;
                                        }
                                    }
                                    else
                                    {
                                        lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: La descripción '{descripcion}' no se encontró en el diccionario.");
                                        //lista_log.Add(Environment.NewLine);
                                        registro_valido = false;
                                    }
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El valor en la celda es inválido (vacío o solo espacios).");
                                    //lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> TipoIdentificador: El campo es nulo.");
                                //lista_log.Add(Environment.NewLine);
                                registro_valido = false;
                            }

                            //
                            g.n_identificador = Convert.ToString(workSheet.Cells[f, 8].Value);
                            //
                            g.tipo_persona = Convert.ToString(workSheet.Cells[f, 9].Value).Substring(0, 1);

                            if (g.tipo_persona == "F")
                            {
                                g.nombre_de_pila = Convert.ToString(workSheet.Cells[f, 11].Value);
                                g.primer_apellido = Convert.ToString(workSheet.Cells[f, 12].Value);
                            }
                            else
                                g.razon_social = Convert.ToString(workSheet.Cells[f, 10].Value);
                            //
                            // Variable para la celda de Teléfono
                            int celdaTelefono = 13;
                            if (workSheet.Cells[f, celdaTelefono].Value != null)
                            {
                                string telefono = Convert.ToString(workSheet.Cells[f, celdaTelefono].Value).Trim();
                                if (telefono.Length >= 6 && telefono.Length <= 12 && telefono.All(char.IsDigit))
                                {
                                    g.telefono = telefono;
                                }
                                else
                                {
                                    lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Teléfono: el campo Teléfono no tiene un formato válido (debe tener entre 6 y 12 dígitos).");
                                    lista_log.Add(Environment.NewLine);
                                    registro_valido = false;
                                }
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Teléfono: el campo Teléfono es nulo.");
                                registro_valido = false;
                            }

                            // Asignar tlf_contacto con el valor de teléfono
                            g.tlf_contacto = g.telefono;

                            // Variable para la celda de Persona Contacto
                            int celdaPersonaContacto = 14;
                            if (workSheet.Cells[f, celdaPersonaContacto].Value != null)
                            {
                                g.persona_contacto = Convert.ToString(workSheet.Cells[f, celdaPersonaContacto].Value).Trim();
                            }
                            else
                            {
                                lista_log.Add($"Hoja: {g.hoja} Fila: {i} --> Persona Contacto: el campo Persona Contacto es nulo.");
                                registro_valido = false;
                            }

                            g.entorno = Convert.ToString(workSheet.Cells[f, 15].Value);
                            g.tipo_cliente = Convert.ToString(workSheet.Cells[f, 16].Value);

                            g.observaciones = Convert.ToString(workSheet.Cells[f, 17].Value);
                            #endregion

                         if (registro_valido)
                         {
                                lista_datos_excel_extrasistemas.Add(g);
                                if (dic_total_registros.TryGetValue(g.hoja, out o))
                                    o.registros = o.registros + 1;
                         }

                        }
                    }

                }
                fs = null;
                excelPackage = null;

                if (dic_total_registros.Count > 0)
                {
                    //lista_log.Add(System.Environment.NewLine);
                    lista_log.Add("Fin de la importación. Puede generar los XML");
                }
                else
                {
                    //lista_log.Add(System.Environment.NewLine);
                    lista_log.Add("Fin de la importación. No se cargado ningún registro.");
                }
                //GuardaDatos_AD(lista_datos_excel_extrasistemas);
                // GuardaDatos_ACC(lista);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + f + " ver--xxxxx-- ",
                      "Extrasistemas.CargaExcel",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Error);

            }
        }

        public void Crea_XML()
        {
            if (lista_datos_excel_extrasistemas.Count > 0)
            {
                DialogResult result = MessageBox.Show("¿Desea generar los archivos XML de los registros validados?",
               "Creador Pasos XML",
            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    GeneraXML(lista_datos_excel_extrasistemas);

                    MessageBox.Show("Proceso completado.",
                      "Creador Pasos XML",
                         MessageBoxButtons.OK, MessageBoxIcon.Information);
                }


            }
            else
            {
                MessageBox.Show("No hay ningún registro validado para poder generar XML.",
                      "Creador Pasos XML",
                         MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
                
        }

        private void GuardaDatos_AD(List<EndesaEntity.extrasistemas.Global> lista_excel)
        {
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            MySQLDB db;
            MySqlCommand command;
            int x = 0;

            try
            {

                List<EndesaEntity.extrasistemas.Global> lista = lista_excel.Where(z => z.hoja == "AD").ToList();

                foreach (EndesaEntity.extrasistemas.Global p in lista)
                {

                    x++;
                    if (firstOnly)
                    {
                        sb.Append("REPLACE INTO cont.extrasistemas_excel_ad");
                        sb.Append(" (hoja, empresa_emisora, distribuidora, cups22,");
                        sb.Append(" cnae, ind_activacion, fecha_activacion,");
                        sb.Append(" tipo_contrato_atr, fecha_finalizacion, tipo_autoconsumo, tarifa,");
                        sb.Append(" p1, p2, p3, p4, p5, p6,");
                        sb.Append(" tipo_identificador, n_identificador, tipo_persona, razon_social, nombre_pila,");
                        sb.Append(" primer_apellido, telefono, persona_contacto, indicador_tipo_direccion, pais, provincia,");
                        sb.Append(" municipio, codigo_postal, tipo_via, nombre_via, numero, entorno, tipo_cliente,");
                        sb.Append(" tipo_doc_aportado, direccion_url, observaciones, xml_generado, fichero,");
                        sb.Append(" created_by, created_date) values ");

                        firstOnly = false;
                    }

                    sb.Append("('").Append(p.hoja).Append("',");
                    sb.Append("'").Append(p.empresa_emisora).Append("',");
                    sb.Append("'").Append(p.distribuidora).Append("',");
                    sb.Append("'").Append(p.cups22).Append("',");                    

                    if (p.cnae != null)
                        sb.Append("'").Append(p.cnae).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.ind_activacion != null)
                        sb.Append("'").Append(p.ind_activacion).Append("',");
                    else
                        sb.Append("null,");

                    if (p.fecha_activacion > DateTime.MinValue)
                        sb.Append("'").Append(p.fecha_activacion.ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.tipo_contrato_atr != null)
                        sb.Append("'").Append(p.tipo_contrato_atr).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.fecha_finalizacion > DateTime.MinValue)
                        sb.Append("'").Append(p.fecha_finalizacion.ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.tipo_autoconsumo != null)
                        sb.Append("'").Append(p.tipo_autoconsumo).Append("',");
                    else
                        sb.Append("null,");                   

                    if (p.tarifa != null)
                        sb.Append("'").Append(p.tarifa).Append("',");
                    else
                        sb.Append("null,");

                    for(int pot = 0; pot < 6; pot++)
                    {
                        if (p.potencias[pot] > 0)
                            sb.Append(p.potencias[pot]).Append(",");
                        else
                            sb.Append(" null,");
                    }

                    if (p.tipo_identificador != null)
                        sb.Append("'").Append(p.tipo_identificador).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.n_identificador != null)
                        sb.Append("'").Append(p.n_identificador).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.tipo_persona != null)
                        sb.Append("'").Append(p.tipo_persona).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.razon_social != null)
                        sb.Append("'").Append(p.razon_social).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.nombre_de_pila != null)
                        sb.Append("'").Append(p.nombre_de_pila).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.primer_apellido != null)
                        sb.Append("'").Append(p.primer_apellido).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.telefono != null)
                        sb.Append("'").Append(p.telefono).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.persona_contacto != null)
                        sb.Append("'").Append(p.persona_contacto).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.indicador_tipo_direccion != null)
                        sb.Append("'").Append(p.indicador_tipo_direccion).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.pais != null)
                        sb.Append("'").Append(p.pais).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.provincia != null)
                        sb.Append("'").Append(p.provincia).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.municipio != null)
                        sb.Append("'").Append(p.municipio).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.codigo_postal != null)
                        sb.Append("'").Append(p.codigo_postal).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.tipo_via != null)
                        sb.Append("'").Append(p.tipo_via).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.nombre_via != null)
                        sb.Append("'").Append(p.nombre_via).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.numero != null)
                        sb.Append("'").Append(p.numero).Append("',");
                    else
                        sb.Append(" null,");
                    // irh    EndesaBusiness- Extrasistemass.cs --  piso  - puerta 
                    //if (p.piso != null)
                    //    sb.Append("'").Append(p.piso).Append("',");
                    //else
                    //    sb.Append(" null,");
                    //
                    //if (p.puerta != null)
                    //    sb.Append("'").Append(p.puerta).Append("',");
                    //else
                    //    sb.Append(" null,");
                    if (p.entorno != null)
                        sb.Append("'").Append(p.entorno).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.tipo_cliente != null)
                        sb.Append("'").Append(p.tipo_cliente).Append("',");
                    else
                        sb.Append(" null,");

                    //if (p.tipo_doc_aportado != null)
                    //    sb.Append("'").Append(p.tipo_doc_aportado).Append("',");
                    //else
                    //    sb.Append(" null,");

                    if (p.direccion_url != null)
                        sb.Append("'").Append(p.direccion_url).Append("',");
                    else
                        sb.Append(" null,");

                    if (p.observaciones != null)
                        sb.Append("'").Append(p.observaciones).Append("',");
                    else
                        sb.Append(" null,");

                    // xml_generado
                    sb.Append("'N',");
                    sb.Append("'").Append(p.fichero).Append("',");
                    sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                    if (x == 100)
                    {                        
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        x = 0;
                    }

                }

                if (x > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    x = 0;
                }


            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message,
                       "GuardaDatos_AD",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Error);
            }
        }
        private void GuardaDatos_ACC(List<EndesaEntity.extrasistemas.Global> lista_excel)
        {
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            MySQLDB db;
            MySqlCommand command;
            int x = 0;

            try
            {

                List<EndesaEntity.extrasistemas.Global> lista = lista_excel.Where(z => z.hoja == "ACC").ToList();

                foreach (EndesaEntity.extrasistemas.Global p in lista)
                {

                    x++;
                    if (firstOnly)
                    {
                        sb.Append("REPLACE INTO cont.extrasistemas_excel_ad");
                        sb.Append(" (hoja, empresa_emisora, distribuidora, cups22,");
                        sb.Append(" ind_activacion, fecha_activacion,");
                        sb.Append(" contrato_incondicional_ps, contratos_incondicional_bs,");                        
                        sb.Append(" tipo_idenficador, n_identificador, tipo_persona, razon_social, nombre_pila,");
                        sb.Append(" primer_apellido, telefono, entorno, tipo_cliente,");
                        sb.Append(" xml_generado, fichero,");
                        sb.Append(" created_by, created_date) values ");

                        firstOnly = false;
                    }

                    sb.Append("('").Append(p.hoja).Append("',");
                    sb.Append("'").Append(p.empresa_emisora).Append("',");
                    sb.Append("'").Append(p.distribuidora).Append("',");
                    sb.Append("'").Append(p.cups22).Append("',");                    

                    if (p.ind_activacion != null)
                        sb.Append("'").Append(p.ind_activacion).Append("',");
                    else
                        sb.Append("null,");

                    if (p.fecha_activacion > DateTime.MinValue)
                        sb.Append("'").Append(p.fecha_activacion.ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (p.contratacion_incondicional_ps != null)
                        sb.Append("'").Append(p.contratacion_incondicional_ps).Append("',");
                    else
                        sb.Append("null,");

                    if (p.contratacion_incondicional_bs != null)
                        sb.Append("'").Append(p.contratacion_incondicional_bs).Append("',");
                    else
                        sb.Append("null,");                    

                    if (p.tipo_identificador != null)
                        sb.Append("'").Append(p.tipo_identificador).Append("',");
                    else
                        sb.Append("null,");
                    
                    if (p.n_identificador != null)
                        sb.Append("'").Append(p.n_identificador).Append("',");
                    else
                        sb.Append("null,");

                    if (p.tipo_persona != null)
                        sb.Append("'").Append(p.tipo_persona).Append("',");
                    else
                        sb.Append("null,");

                    if (p.razon_social != null)
                        sb.Append("'").Append(p.razon_social).Append("',");
                    else
                        sb.Append("null,");

                    if (p.nombre_de_pila != null)
                        sb.Append("'").Append(p.nombre_de_pila).Append("',");
                    else
                        sb.Append("null,");

                    if (p.primer_apellido != null)
                        sb.Append("'").Append(p.primer_apellido).Append("',");
                    else
                        sb.Append("null,");

                    if (p.telefono != null)
                        sb.Append("'").Append(p.telefono).Append("',");
                    else
                        sb.Append("null,");

                    

                    if (p.entorno != null)
                        sb.Append("'").Append(p.entorno).Append("',");
                    else
                        sb.Append("null,");

                    if (p.tipo_cliente != null)
                        sb.Append("'").Append(p.tipo_cliente).Append("',");
                    else
                        sb.Append("null,");

                    if (p.observaciones != null)
                        sb.Append("'").Append(p.observaciones).Append("',");
                    else
                        sb.Append("null,");

                    // xml_generado
                    sb.Append("'N',");
                    sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                    if (x == 100)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        x = 0;
                    }

                }

                if (x > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    x = 0;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                       "GuardaDatos_AD",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Error);
            }
        }

        private void GeneraXML(List<EndesaEntity.extrasistemas.Global> lista)
        {
            XmlWriterSettings settings;
            XmlWriter writer;
            XmlSerializer serializer;
            FileInfo file;
            
            string mensaje_error = "";

            try
            {

                foreach(EndesaEntity.extrasistemas.Global p in lista)
                {
                    if (p.hoja == "AD")    // xml
                    {
                        //NO PROCESAMOS ESTA PESTAÑA HASTA FINALIZAR IMPLEMENTACION
                        //FORZAMOS BREAK
                        //lista_log.Add("La pestaña AD no se procesará, está deshabilitada temporalmente");
                        //continue;
                        //

                        file = new FileInfo(param.GetValue("ruta_salida_xml") + "A301"
                            + "_" + p.cups22 + "_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xml");

                        EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 a301 =
                            new EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301();

                       
                        #region Cabecera

                        //EndesaEntity.cnmc.V21_2019_12_17.Cabecera cabecera = new EndesaEntity.cnmc.V21_2019_12_17.Cabecera();
                        
                        //cabecera.CodigoREEEmpresaEmisora = p.empresa_emisora;
                        //cabecera.CodigoREEEmpresaDestino = p.distribuidora;
                        //cabecera.CodigoDelProceso = "A3";
                        //cabecera.CodigoDePaso = "01";
                        //cabecera.CodigoDeSolicitud = DateTime.Now.ToString("yyMMddHHmmss");
                        //cabecera.SecuencialDeSolicitud = "00";
                        //cabecera.FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
                        //cabecera.CUPS = p.cups22;

                        //a301.Cabecera = cabecera;

                        a301.Cabecera.CodigoREEEmpresaEmisora = p.empresa_emisora;
                        a301.Cabecera.CodigoREEEmpresaDestino = p.distribuidora;
                        a301.Cabecera.CodigoDelProceso = "A3";
                        a301.Cabecera.CodigoDePaso = "01";
                        a301.Cabecera.CodigoDeSolicitud = DateTime.Now.ToString("yyMMddHHmmss");
                        a301.Cabecera.SecuencialDeSolicitud = "00";
                        a301.Cabecera.FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
                        a301.Cabecera.CUPS = p.cups22;

                        #endregion

                        #region Alta

                        #region DatosSolicitud

                        a301.Alta.DatosSolicitud.cnae = p.cnae;
                        a301.Alta.DatosSolicitud.IndActivacion = p.ind_activacion;

                        if(p.ind_activacion == "F")
                            a301.Alta.DatosSolicitud.fechaPrevistaAccion = p.fecha_activacion.ToString("yyyy-MM-dd");

                        //a301.Alta.DatosSolicitud.SolicitudTension = "N";
                        
                        #endregion

                        #region Contrato
                        {

                            a301.Alta.Contrato.FechaFinalizacion = p.fecha_finalizacion.ToString("yyyy-MM-dd");

                            #region Autoconsumo

                          
                            if (p.tipo_autoconsumo != "00" && p.tipo_autoconsumo != "0C")
                            {
                               a301.Alta.Contrato.Autoconsumo = new AutoconsumoSolicitudAlta();

                                a301.Alta.Contrato.Autoconsumo.DatosSuministro.TipoCUPS = p.tipocups;
                                a301.Alta.Contrato.Autoconsumo.DatosCAU.CAU = p.cau;
                                a301.Alta.Contrato.Autoconsumo.DatosCAU.TipoAutoconsumo = p.tipo_autoconsumo;
                                a301.Alta.Contrato.Autoconsumo.DatosCAU.TipoSubseccion = p.tipo_subseccion;
                                a301.Alta.Contrato.Autoconsumo.DatosCAU.Colectivo = p.colectivo;
                                a301.Alta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.PotInstaladaGen = p.potinstaladagen.ToString();
                                a301.Alta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.TipoInstalacion = p.tipoinstalacion;
                                a301.Alta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.SSAA = p.ssaa;
                                a301.Alta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.UnicoContrato = p.unicocontrato;
                             
                            }
                            //a301.Alta.Contrato.Autoconsumo.TipoAutoconsumo = p.tipo_autoconsumo;
                            #endregion

                            
                            a301.Alta.Contrato.TipoContratoATR = p.tipo_contrato_atr;

                            // a301.Alta.Contrato.CUPSPrincipal
                            if (p.tipo_contrato_atr == "08" || p.tipo_contrato_atr == "11" || p.tipo_contrato_atr == "12")
                            {
                                a301.Alta.Contrato.CUPSPrincipal = p.cups22;
                            }
                            // 
                            // 
                            {
                                #region CondicionesContractuales


                                a301.Alta.Contrato.CondicionesContractuales.TarifaATR = p.tarifa.Substring(0,3);
                                PotenciasContratadas potenciasContratadas = new PotenciasContratadas();
                                

                                for (int i = 1; i <= 6; i++)
                                {
                                    if(p.potencias[i-1] != 0)
                                    {
                                        Potencia potencia = new Potencia();
                                        potencia.periodo = Convert.ToString(i);
                                        potencia.potencia = Convert.ToString(p.potencias[i-1]);
                                        potenciasContratadas.Potencia.Add(potencia);
                                    }
                                        
                                }
                                a301.Alta.Contrato.CondicionesContractuales.PotenciasContratadas = potenciasContratadas;
                                #endregion


                                #region Contacto
                                //EndesaEntity.cnmc.V21_2019_12_17.Contacto contacto = new Contacto();
                                a301.Alta.Contrato.Contacto.PersonaDeContacto = p.persona_contacto;
                                //a301.Alta.Contrato.Contacto = contacto;

                                
                                a301.Alta.Contrato.Contacto.Telefono.PrefijoPais = "0034";
                                a301.Alta.Contrato.Contacto.Telefono.Numero = p.tlf_contacto;

                                #endregion

                            }
                        }
                        #endregion

                        #region Cliente
                        {
                            //a301.Alta.Cliente.IdCliente.TipoIdentificador = cnmc.GetTipo_Identificador(p.tipo_identificador);
                            a301.Alta.Cliente.IdCliente.TipoIdentificador = p.tipo_identificador;
                            a301.Alta.Cliente.IdCliente.Identificador = p.n_identificador;
                            a301.Alta.Cliente.IdCliente.TipoPersona = p.tipo_persona.Substring(0,1);

                            if (a301.Alta.Cliente.IdCliente.TipoPersona == "J")
                                a301.Alta.Cliente.Nombre.RazonSocial = p.razon_social;
                            else
                            {
                                a301.Alta.Cliente.Nombre.NombreDePila = p.nombre_de_pila;
                                a301.Alta.Cliente.Nombre.PrimerApellido = p.primer_apellido;
                            }

                            a301.Alta.Cliente.Telefono.PrefijoPais = "0034";
                            a301.Alta.Cliente.Telefono.Numero = p.telefono;


                            a301.Alta.Cliente.IndicadorTipoDireccion = p.indicador_tipo_direccion.Substring(0, 1);
                            a301.Alta.Cliente.Direccion.Pais = p.pais;
                            a301.Alta.Cliente.Direccion.Provincia = p.provincia.Substring(0, 2);
                            //a301.Alta.Cliente.Direccion.Municipio =
                            //   cnmc.GetCodigoMunicipio(a301.Alta.Cliente.Direccion.Provincia, p.municipio);
                            a301.Alta.Cliente.Direccion.Municipio = p.municipio;                            
                            a301.Alta.Cliente.Direccion.CodPostal = p.codigo_postal;
                            a301.Alta.Cliente.Direccion.Via.TipoVia = p.tipo_via;
                            a301.Alta.Cliente.Direccion.Via.Calle = p.nombre_via;
                            a301.Alta.Cliente.Direccion.Via.NumeroFinca = p.numero;

                            // irh  19/03/25
                            if (!string.IsNullOrEmpty(p.piso))
                            {
                                a301.Alta.Cliente.Direccion.Via.Piso = p.piso;
                            }
                            //a301.Alta.Cliente.Direccion.Via.Piso = p.piso;
                            if (!string.IsNullOrEmpty(p.puerta))
                            {
                                a301.Alta.Cliente.Direccion.Via.Puerta = p.puerta;
                            }//  a301.Alta.Cliente.Direccion.Via.Puerta = p.puerta;

                        }
                        #endregion

                        #region RegistroDocumento
                        if (p.lista_documentacion.Count > 0)
                        {
                            a301.Alta.RegistrosDocumento = new RegistrosDocumento();

                            foreach (Documentacion pp in p.lista_documentacion)
                            {
                                RegistroDoc regdoc = new RegistroDoc();
                                regdoc.TipoDocAportado = pp.tipo_doc_aportado;
                                regdoc.DireccionUrl = pp.direccion_url;
                                a301.Alta.RegistrosDocumento.RegistroDoc.Add(regdoc);
                            }
                        }

                        if (!System.IO.Directory.Exists(param.GetValue("ruta_salida_xml")))
                        {
                            System.IO.Directory.CreateDirectory(param.GetValue("ruta_salida_xml"));
                        }
                        #endregion

                        #endregion

                        settings = new XmlWriterSettings();
                        settings.Indent = true;
                        settings.Encoding = Encoding.UTF8;
                        settings.OmitXmlDeclaration = true;
                        writer = XmlWriter.Create(file.FullName, settings);

                        XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                        ns.Add("", @"http://localhost/elegibilidad");
                        serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301));

                        serializer.Serialize(writer, a301, ns);
                        writer.Close();

                        mensaje_error = ValidateSchema(file.FullName, 
                            System.Environment.CurrentDirectory + param_cnmc.GetValue("xsd_a301"));

                        if (mensaje_error != "")
                        {
                            //lista_log.Add(System.Environment.NewLine);
                            lista_log.Add("Hoja: " + p.hoja + " CUPS: " + p.cups22 +
                                        " --> " + mensaje_error);
                            //
                            file.Delete();
                        }else
                        {
                            //lista_log.Add(System.Environment.NewLine);
                            lista_log.Add("Hoja: " + p.hoja + " CUPS: " + p.cups22 +
                                        " --> " + "XML generado correctamente");
                        }
               
                    }

                    if (p.hoja == "ACC")
                    {

                        //NO PROCESAMOS ESTA PESTAÑA HASTA FINALIZAR IMPLEMENTACION  -------   xml  -----
                        //FORZAMOS BREAK
                        //lista_log.Add("La pestaña ACC no se procesará, está deshabilitada temporalmente");
                        //continue;
                        //

                        file = new FileInfo(param.GetValue("ruta_salida_xml") + "C101"
                            + "_" + p.cups22 + "_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xml");

                        EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 c101 = new EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101();
                        #region Cabecera

                        EndesaEntity.cnmc.V21_2019_12_17.Cabecera cabecera = new EndesaEntity.cnmc.V21_2019_12_17.Cabecera();

                        cabecera.CodigoREEEmpresaEmisora = p.empresa_emisora;

                        // 27/06/2024 GUS: modificamos porque recibimos directamente el código de la distribuidora
                        //cabecera.CodigoREEEmpresaDestino = cnmc.Distribuidora(p.distribuidora);
                        cabecera.CodigoREEEmpresaDestino = p.distribuidora;
                        cabecera.CodigoDelProceso = "C1";
                        cabecera.CodigoDePaso = "01";
                        cabecera.CodigoDeSolicitud = DateTime.Now.ToString("yyMMddHHmmss");
                        cabecera.SecuencialDeSolicitud = "00";
                        cabecera.FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
                        cabecera.CUPS = p.cups22;
                        c101.Cabecera = cabecera;
                        #endregion

                        #region DatosSolicitud
                        //
                        c101.CambiodeComercializadorSinCambios.DatosSolicitud.indActivacion = p.ind_activacion;

                        // if (c101.CambiodeComercializadorSinCambios.DatosSolicitud.indActivacion == "F")
                        if (p.ind_activacion == "F")
                            c101.CambiodeComercializadorSinCambios.DatosSolicitud.fechaPrevistaAccion = p.fecha_activacion.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                       
                        c101.CambiodeComercializadorSinCambios.DatosSolicitud.contratacionIncondicionalPS = p.contratacion_incondicional_ps;
                          
                        // Siempre N
                        c101.CambiodeComercializadorSinCambios.DatosSolicitud.contratacionIncondicionalBS = "N";

                        #endregion

                        #region Cliente

                        EndesaEntity.cnmc.V21_2019_12_17.Cliente cliente = new EndesaEntity.cnmc.V21_2019_12_17.Cliente();

                       //cliente.IdCliente.TipoIdentificador =  cnmc.GetTipo_Identificador(p.tipo_identificador);
                        cliente.IdCliente.TipoIdentificador = p.tipo_identificador;

                        cliente.IdCliente.Identificador = p.n_identificador;

                        // Solo se deberá informar este campo cuando el TipoIdentificador sea "NIVA",  "Otros" -  irh 19/03/25

                        if (p.tipo_identificador == "NV" || p.tipo_identificador == "OT")
                        {
                            cliente.IdCliente.TipoPersona = p.tipo_persona;
                        }
                        //cliente.IdCliente.TipoPersona = p.tipo_persona;

                        c101.CambiodeComercializadorSinCambios.Cliente = cliente;
                        #endregion

                        #region Nombre
                        
                        c101.CambiodeComercializadorSinCambios.Cliente.Nombre.NombreDePila =
                                p.nombre_de_pila;
                        c101.CambiodeComercializadorSinCambios.Cliente.Nombre.PrimerApellido =
                                p.primer_apellido;
                        
                        c101.CambiodeComercializadorSinCambios.Cliente.Nombre.RazonSocial =
                                p.razon_social;
                        
                        #endregion

                        c101.CambiodeComercializadorSinCambios.Cliente.Telefono.PrefijoPais = "0034";
                        c101.CambiodeComercializadorSinCambios.Cliente.Telefono.Numero = p.telefono;
                        c101.CambiodeComercializadorSinCambios.Cliente.Telefono.CorreoElectronico = p.contacto;

                        foreach (Documentacion pp in p.lista_documentacion)
                        {
                            RegistroDoc regdoc = new RegistroDoc();
                            regdoc.TipoDocAportado = pp.tipo_doc_aportado;
                            regdoc.DireccionUrl = pp.direccion_url;
                            c101.CambiodeComercializadorSinCambios.RegistrosDocumento.RegistroDoc.Add(regdoc);
                        }
                        
                        if (!System.IO.Directory.Exists(param.GetValue("ruta_salida_xml")))
                        {
                            System.IO.Directory.CreateDirectory(param.GetValue("ruta_salida_xml"));
                        }


                        settings = new XmlWriterSettings();
                        settings.Indent = true;
                        settings.Encoding = Encoding.UTF8;
                        settings.OmitXmlDeclaration = true;
                        writer = XmlWriter.Create(file.FullName, settings);

                        XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                        ns.Add("", "http://localhost/elegibilidad");
                        serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101));
                       
                        serializer.Serialize(writer, c101, ns);
                        writer.Close();

                        mensaje_error = ValidateSchema(file.FullName, System.Environment.CurrentDirectory + param_cnmc.GetValue("xsd_c101"));
                        // irh
                        if (mensaje_error != "")
                        {
                            //lista_log.Add(System.Environment.NewLine);
                            lista_log.Add("Hoja: " + p.hoja + " CUPS: " + p.cups22 +
                                        " --> " + mensaje_error);
                            file.Delete();
                        }
                        else
                        {
                            //lista_log.Add(System.Environment.NewLine);
                            lista_log.Add("Hoja: " + p.hoja + " CUPS: " + p.cups22 +
                                        " --> " + "XML generado correctamente");
                        }

                    }

                    if (p.hoja == "ACC + CAMBIOS")  //  xml
                    {

                        //NO PROCESAMOS ESTA PESTAÑA HASTA FINALIZAR IMPLEMENTACION
                        //FORZAMOS BREAK
                       // lista_log.Add("La pestaña Acc + cambios --  no se procesará, está deshabilitada temporalmente");
                       // continue;

                        file = new FileInfo(param.GetValue("ruta_salida_xml") + "C201"
                            + "_" + p.cups22 + "_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xml");

                        EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeC201 c201 = new EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeC201();

                        //EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeB101 b101 =
                        //    new EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeB101();
                #region Cabecera

                        EndesaEntity.cnmc.V21_2019_12_17.Cabecera cabecera = new EndesaEntity.cnmc.V21_2019_12_17.Cabecera();

                        cabecera.CodigoREEEmpresaEmisora = p.empresa_emisora;

                        cabecera.CodigoREEEmpresaDestino = p.distribuidora; //cnmc.Distribuidora(p.distribuidora);

                        cabecera.CodigoDelProceso = "C2";
                        cabecera.CodigoDePaso = "01";
                        cabecera.CodigoDeSolicitud = DateTime.Now.ToString("yyMMddHHmmss");
                        cabecera.SecuencialDeSolicitud = "00";
                        cabecera.FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
                        cabecera.CUPS = p.cups22;
                        c201.Cabecera = cabecera;
                #endregion
               //CambiodeComercialzadorConCambios --
                   #region DatosSolicitud
                        //tipo_modificacion -------- DatosSolicitud_C1 ---------
                        c201.CambioComercializadorConCambios.DatosSolicitud.tipoModificacion = p.tipo_modificacion;
                        // tipo solicitud administrativa
                        c201.CambioComercializadorConCambios.DatosSolicitud.tipoSolicitudAdministrativa = p.tipo_solicitud_administrativa;
                        // cnae 
                        c201.CambioComercializadorConCambios.DatosSolicitud.cnae = p.cnae;
                        // ind activacion 
                        c201.CambioComercializadorConCambios.DatosSolicitud.indActivacion = p.ind_activacion;
                        // fecha activacion
                        if (p.ind_activacion == "F")
                        {
                            c201.CambioComercializadorConCambios.DatosSolicitud.fechaPrevistaAccion = p.fecha_activacion.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                        }
                        //contratacion incondicional PS 
                        c201.CambioComercializadorConCambios.DatosSolicitud.contratacionIncondicionalPS = p.contratacion_incondicional_ps;
                        // Contratacion incond BS 
                        c201.CambioComercializadorConCambios.DatosSolicitud.contratacionIncondicionalBS = p.contratacion_incondicional_bs;
                        // solicitud modificacion tension -------
                        c201.CambioComercializadorConCambios.DatosSolicitud.solicitudTension = p.solicitud_modificacion_tension;
                        //
                        if (p.solicitud_modificacion_tension == "S")
                        {
                            c201.CambioComercializadorConCambios.DatosSolicitud.TensionSolicitada = p.tension_solicitada;
                        }
                #endregion
                   
                        //------------  contrato -------
                        //   contratos obligatorio informar este nodo si "TipoModificacion" es igual a "A" o "N".
                        if (p.tipo_modificacion == "A" || p.tipo_modificacion == "N")
                        {
                            // inicializar contrato.
                          c201.CambioComercializadorConCambios.Contrato = new Contrato_C201();

                            // fecha de finalizacion 
                            if (p.tipo_contrato_atr == "02" || p.tipo_contrato_atr == "03" || p.tipo_contrato_atr == "09")
                            {
                                c201.CambioComercializadorConCambios.Contrato.FechaFinalizacion =
                                                     p.fecha_finalizacion.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                            }

                            // Autoconsumo ----------------
                            // Datos Suministro
                            // inicilizar autoconsumo - DatosSuministroSolicitud - DatosCAUAlta
                            bool debeInformarCamposDS = p.tipo_autoconsumo != "00" && p.tipo_autoconsumo != "0C";
                            if (debeInformarCamposDS)
                            {

                                c201.CambioComercializadorConCambios.Contrato.Autoconsumo = new AutoconsumoSolicitudAlta();
                                c201.CambioComercializadorConCambios.Contrato.Autoconsumo.DatosSuministro.TipoCUPS = p.tipocups;

                            }
                            //inicializo datos cau
                            c201.CambioComercializadorConCambios.Contrato.Autoconsumo.DatosCAU = new DatosCAUAlta();
                            //Datos CAU
                            c201.CambioComercializadorConCambios.Contrato.Autoconsumo.DatosCAU.CAU = p.cau;
                            // tipo de autoconsumo 
                            c201.CambioComercializadorConCambios.Contrato.Autoconsumo.DatosCAU.TipoAutoconsumo = p.tipo_autoconsumo;
                            //TipoSubseccion  
                            c201.CambioComercializadorConCambios.Contrato.Autoconsumo.DatosCAU.TipoSubseccion = p.tipo_subseccion;
                            //
                            if (p.tipo_autoconsumo != "00" && p.tipo_autoconsumo != "0C")
                            {
                                c201.CambioComercializadorConCambios.Contrato.Autoconsumo.DatosCAU.Colectivo = p.colectivo;
                            }
                            //----- DatosInstGen  ------------------------
                            bool debeInformarCamposIG = p.tipo_autoconsumo != "00" && p.tipo_autoconsumo != "0C";
                            if (debeInformarCamposIG)
                            {
                                c201.CambioComercializadorConCambios.Contrato.Autoconsumo.DatosCAU.DatosInstGen.PotInstaladaGen = p.potinstaladagen.ToString();

                                //if (p.tipo_autoconsumo == "11")
                                // {
                                c201.CambioComercializadorConCambios.Contrato.Autoconsumo.DatosCAU.DatosInstGen.TipoInstalacion = p.tipoinstalacion;
                                //}
                                c201.CambioComercializadorConCambios.Contrato.Autoconsumo.DatosCAU.DatosInstGen.SSAA = p.ssaa;
                                c201.CambioComercializadorConCambios.Contrato.Autoconsumo.DatosCAU.DatosInstGen.UnicoContrato = p.unicocontrato;

                            }
                            //TipoContratoATR 
                            c201.CambioComercializadorConCambios.Contrato.TipoContratoATR = p.tipo_contrato_atr;
                            // Alta.Contrato.CUPSPrincipal
                            if (p.tipo_contrato_atr == "08" || p.tipo_contrato_atr == "11" || p.tipo_contrato_atr == "12")
                            {
                                c201.CambioComercializadorConCambios.Contrato.CUPSPrincipal = p.cups22;
                            }
                            //
                            //    condicionescontractuales
                            //------------------- tarifa   potencia 
                            // inicializo
                            c201.CambioComercializadorConCambios.Contrato.CondicionesContractuales = new CondicionesContractuales_C201();
                                
                            c201.CambioComercializadorConCambios.Contrato.CondicionesContractuales.TarifaATR = p.tarifa;

                            PotenciasContratadas potenciasContratadas = new PotenciasContratadas();

                            for (int i = 1; i <= 6; i++)
                            {
                                if (p.potencias[i - 1] != 0)
                                {
                                    Potencia potencia = new Potencia();
                                    potencia.periodo = Convert.ToString(i);
                                    potencia.potencia = Convert.ToString(p.potencias[i - 1]);
                                    potenciasContratadas.Potencia.Add(potencia);
                                }
                            }
                            c201.CambioComercializadorConCambios.Contrato.CondicionesContractuales.PotenciasContratadas = potenciasContratadas;
                            //Contacto--------------------
                            // persona de contacto - inicializar
                            c201.CambioComercializadorConCambios.Contrato.Contacto = new Contacto();
                            c201.CambioComercializadorConCambios.Contrato.Contacto.PersonaDeContacto = p.persona_contacto;
                            // telefono 
                            c201.CambioComercializadorConCambios.Contrato.Contacto.Telefono.PrefijoPais = "0034";
                            c201.CambioComercializadorConCambios.Contrato.Contacto.Telefono.Numero = p.tlf_contracto;
                        }
                        //---Cliente ------
                        // inicializar cliente 
                        EndesaEntity.cnmc.V21_2019_12_17.Cliente cliente = new EndesaEntity.cnmc.V21_2019_12_17.Cliente();
                        c201.CambioComercializadorConCambios.Cliente = cliente;
                        // tipo identificador 
                        c201.CambioComercializadorConCambios.Cliente.IdCliente.TipoIdentificador = p.tipo_identificador;
                        // n  identificador 
                        c201.CambioComercializadorConCambios.Cliente.IdCliente.Identificador = p.n_identificador;
                        // tipo persona     
                        c201.CambioComercializadorConCambios.Cliente.IdCliente.TipoPersona = p.tipo_persona;
                        //
                        c201.CambioComercializadorConCambios.Cliente.Telefono.PrefijoPais = "0034";
                        c201.CambioComercializadorConCambios.Cliente.Telefono.Numero = p.telefono;
                                               
                        // razon social  -  nombre de pila -  primer apellido 
                        if (c201.CambioComercializadorConCambios.Cliente.IdCliente.TipoPersona == "J")
                              c201.CambioComercializadorConCambios.Cliente.Nombre.RazonSocial = p.razon_social;
                        {
                             c201.CambioComercializadorConCambios.Cliente.Nombre.NombreDePila = p.nombre_de_pila;
                             c201.CambioComercializadorConCambios.Cliente.Nombre.PrimerApellido = p.primer_apellido;
                        }
                                            
                        // indicar tipo direccion 
                        c201.CambioComercializadorConCambios.Cliente.IndicadorTipoDireccion = p.indicador_tipo_direccion;
                        EndesaEntity.cnmc.V21_2019_12_17.Direccion direccion = new EndesaEntity.cnmc.V21_2019_12_17.Direccion();
                        c201.CambioComercializadorConCambios.Cliente.Direccion = direccion;

                        // pais
                        c201.CambioComercializadorConCambios.Cliente.Direccion.Pais = p.pais;
                        //  provincia 
                        c201.CambioComercializadorConCambios.Cliente.Direccion.Provincia = p.provincia;
                        // municipio
                        c201.CambioComercializadorConCambios.Cliente.Direccion.Municipio = p.municipio;
                        // codigo postal 
                        c201.CambioComercializadorConCambios.Cliente.Direccion.CodPostal = p.codigo_postal;
                        // tipo de via 
                        c201.CambioComercializadorConCambios.Cliente.Direccion.Via.TipoVia = p.tipo_via;
                        // nombre via
                        c201.CambioComercializadorConCambios.Cliente.Direccion.Via.Calle = p.nombre_via;
                        
                        //numero 
                        c201.CambioComercializadorConCambios.Cliente.Direccion.Via.NumeroFinca = p.numero;
                        //entorno  informacion operaciones
                        //tipo cliente  informacion operaciones
                        
                        // tipo doc aportado - direccion url 
                        if (p.lista_documentacion.Count > 0)
                        {
                          c201.CambioComercializadorConCambios.RegistrosDocumento = new RegistrosDocumento();

                            foreach (Documentacion pp in p.lista_documentacion)
                            {
                                RegistroDoc regdoc = new RegistroDoc();
                                regdoc.TipoDocAportado = pp.tipo_doc_aportado;
                                regdoc.DireccionUrl = pp.direccion_url;
                                c201.CambioComercializadorConCambios.RegistrosDocumento.RegistroDoc.Add(regdoc);    
                            }
                        }
                        //

                        if (!System.IO.Directory.Exists(param.GetValue("ruta_salida_xml")))
                        {
                            System.IO.Directory.CreateDirectory(param.GetValue("ruta_salida_xml"));
                        }

                        settings = new XmlWriterSettings();
                        settings.Indent = true;
                        settings.Encoding = Encoding.UTF8;
                        settings.OmitXmlDeclaration = true;
                        writer = XmlWriter.Create(file.FullName, settings);

                        XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                        ns.Add("", @"http://localhost/elegibilidad");

                        serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeC201));

                        serializer.Serialize(writer, c201, ns);
                        writer.Close();

                        mensaje_error = ValidateSchema(file.FullName,
                            System.Environment.CurrentDirectory + param_cnmc.GetValue("xsd_c201"));

                        if (mensaje_error != "")
                        {
                           // lista_log.Add(System.Environment.NewLine);
                            lista_log.Add("Hoja: " + p.hoja + " CUPS: " + p.cups22 +
                                        " --> " + mensaje_error);
                            //file.Delete();

                        }
                        else
                        {
                            //lista_log.Add(System.Environment.NewLine);
                            lista_log.Add("Hoja: " + p.hoja + " CUPS: " + p.cups22 +
                                        " --> " + "XML generado correctamente");
                        }

                    }  //  xml

                    if (p.hoja == "MOD")  // xml
                    {

                        //NO PROCESAMOS ESTA PESTAÑA HASTA FINALIZAR IMPLEMENTACION
                        //FORZAMOS BREAK
                        //lista_log.Add("La pestaña MOD no se procesará, está deshabilitada temporalmente");
                        //continue;

                        file = new FileInfo(param.GetValue("ruta_salida_xml") + "M101"
                            + "_" + p.cups22 + "_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xml");

                        EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101 m101 =
                            new EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101();

                        #region Cabecera

                        EndesaEntity.cnmc.V21_2019_12_17.Cabecera cabecera =
                            new EndesaEntity.cnmc.V21_2019_12_17.Cabecera();

                        cabecera.CodigoREEEmpresaEmisora = p.empresa_emisora;
                        cabecera.CodigoREEEmpresaDestino = p.distribuidora;
                        cabecera.CodigoDelProceso = "M1";
                        cabecera.CodigoDePaso = "01";
                        cabecera.CodigoDeSolicitud = DateTime.Now.ToString("yyMMddHHmmss");
                        cabecera.SecuencialDeSolicitud = "00";
                        cabecera.FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
                        cabecera.CUPS = p.cups22;
                        m101.Cabecera = cabecera;
                        #endregion

                        DatosSolicitud_M1 datosSolicitud = new DatosSolicitud_M1();

                        Contrato_M1 contratoM = new Contrato_M1();  //  irh 

                        // DatosSolicitud
                        datosSolicitud.TipoModificacion = p.tipo_modificacion;

                        if (p.tipo_modificacion == "S" || p.tipo_modificacion == "A")
                            datosSolicitud.TipoSolicitudAdministrativa = p.tipo_solicitud_administrativa;

                        if (p.cnae != null)
                            datosSolicitud.CNAE = p.cnae;

                        // datosSolicitud.IndEsencial = "00";  // irh -
                        datosSolicitud.IndEsencial = p.indesencial;
                        //
                        datosSolicitud.IndActivacion = p.ind_activacion;  ///  varidicaR Valor . 


                        if (p.ind_activacion == "F")
                            datosSolicitud.FechaPrevistaAccion = p.fecha_activacion.ToString("yyyy-MM-dd");

                        // Validar si el tipo de contrato requiere fecha de finalización
                        if (p.tipo_contrato_atr == "02" || p.tipo_contrato_atr == "03" || p.tipo_contrato_atr == "09" || p.tipo_contrato_atr == "10")
                        {
                            
                            m101.ModificacionDeATR.Contrato.FechaFinalizacion = p.fecha_finalizacion.ToString("yyyy-MM-dd");
                        }

                        // Autoconsumo  -  informacion sobre autoconsumo 
                        //   DatosSuministro 

                        //if (p.tipo_autoconsumo != "00" && p.tipo_autoconsumo != "0C")   //irh  02/05/25
                        //{
                        m101.ModificacionDeATR.Contrato.Autoconsumo.DatosSuministro.TipoCUPS = p.tipocups;
                        //}

                        // DatosSuministroSolicitud

                        //DatosCau   irh
                        m101.ModificacionDeATR.Contrato.Autoconsumo.DatosCAU.CAU = p.cau;

                      
                        m101.ModificacionDeATR.Contrato.Autoconsumo.DatosCAU.TipoAutoconsumo = p.tipo_autoconsumo;

                        //
                        m101.ModificacionDeATR.Contrato.Autoconsumo.DatosCAU.TipoSubseccion = p.tipo_subseccion;

                        //DatosInstGen
                        //if (p.tipo_autoconsumo != "00" && p.tipo_autoconsumo != "0C")  //irh 02/05/25
                        //{
                        // PotInstaladaGen                                     
                        m101.ModificacionDeATR.Contrato.Autoconsumo.DatosCAU.DatosInstGen.PotInstaladaGen = p.potinstaladagen.ToString(); ;
                                                        
                        //}
                        // TipoInstalacion


                        m101.ModificacionDeATR.DatosSolicitud = datosSolicitud;

                        //contratoM.TipoContratoATR = p.tipo_contrato_atr;      //irh
                        m101.ModificacionDeATR.Contrato.TipoContratoATR = p.tipo_contrato_atr;
                                                
                        // CUPSPrincipal
                        if (p.tipo_contrato_atr == "08" || p.tipo_contrato_atr == "11" || p.tipo_contrato_atr == "12")
                        {
                            // contratoM.CUPSPrincipal = p.cups22;    //irh                     
                            m101.ModificacionDeATR.Contrato.CUPSPrincipal = p.cups22;
                           
                        }
                        //CondicionesContractuales

                        if (p.tarifa != null)
                            m101.ModificacionDeATR.Contrato.CondicionesContractuales.TarifaATR = p.tarifa.Substring(0, 3);

                        PotenciasContratadas potenciasContratadas = new PotenciasContratadas();

                        for (int i = 1; i <= 6; i++)
                        {
                            if (p.potencias[i - 1] != 0)
                            {
                                Potencia potencia = new Potencia();
                                potencia.periodo = Convert.ToString(i);
                                potencia.potencia = Convert.ToString(p.potencias[i - 1]);
                                potenciasContratadas.Potencia.Add(potencia);
                            }

                        }
                        m101.ModificacionDeATR.Contrato.CondicionesContractuales.PotenciasContratadas = potenciasContratadas;

                        //
                        //  requerido en el xsl  ModoControlPotencia
                        if (p.modo_control_potencia != null)
                            m101.ModificacionDeATR.Contrato.CondicionesContractuales.ModoControlPotencia = p.modo_control_potencia.Substring(0, 1);

                        //Contacto
                        if (p.contacto != null)
                        {
                            EndesaEntity.cnmc.V21_2019_12_17.Contacto contacto = new Contacto();
                            contacto.PersonaDeContacto = p.persona_contacto;

                            Telefono telefono_contacto = new Telefono();
                            telefono_contacto.PrefijoPais = "0034";
                            telefono_contacto.Numero = p.telefono;

                            contacto.Telefono = telefono_contacto;

                            m101.ModificacionDeATR.Contrato.Contacto = contacto;
                        }
                                               
                        m101.ModificacionDeATR.Cliente.IdCliente.TipoIdentificador = p.tipo_identificador;  

                        m101.ModificacionDeATR.Cliente.IdCliente.Identificador = p.n_identificador;

                        m101.ModificacionDeATR.Cliente.IdCliente.TipoPersona = p.tipo_persona.Substring(0, 1);

                        if (m101.ModificacionDeATR.Cliente.IdCliente.TipoPersona == "J")
                            m101.ModificacionDeATR.Cliente.Nombre.RazonSocial = p.razon_social;
                        else
                        {
                            m101.ModificacionDeATR.Cliente.Nombre.NombreDePila = p.nombre_de_pila;
                            m101.ModificacionDeATR.Cliente.Nombre.PrimerApellido = p.primer_apellido;
                        }

                        m101.ModificacionDeATR.Contrato.Contacto.PersonaDeContacto = p.persona_contacto;   

                        // Siempre se pone el telf de contacto
                        m101.ModificacionDeATR.Cliente.Telefono.PrefijoPais = "0034";
                        m101.ModificacionDeATR.Cliente.Telefono.Numero = p.telefono;

                        m101.ModificacionDeATR.Cliente.IndicadorTipoDireccion = p.indicador_tipo_direccion.Substring(0, 1);

                        
                        m101.ModificacionDeATR.Cliente.Direccion = new Direccion();
                        m101.ModificacionDeATR.Cliente.Direccion.Pais = p.pais; 

                        m101.ModificacionDeATR.Cliente.Direccion.Provincia = p.provincia.Substring(0, 2);

                        m101.ModificacionDeATR.Cliente.Direccion.Municipio = p.municipio;  
                                                                                              
                        m101.ModificacionDeATR.Cliente.Direccion.CodPostal = p.codigo_postal;
                        m101.ModificacionDeATR.Cliente.Direccion.Via.TipoVia = p.tipo_via;
                        m101.ModificacionDeATR.Cliente.Direccion.Via.Calle = p.nombre_via;
                        m101.ModificacionDeATR.Cliente.Direccion.Via.NumeroFinca = p.numero;

                        if (p.lista_documentacion.Count > 0)
                        {
                            m101.ModificacionDeATR.RegistrosDocumento = new RegistrosDocumento();

                            foreach (Documentacion pp in p.lista_documentacion)
                            {
                                RegistroDoc regdoc = new RegistroDoc();
                                regdoc.TipoDocAportado = pp.tipo_doc_aportado;
                                regdoc.DireccionUrl = pp.direccion_url;
                                m101.ModificacionDeATR.RegistrosDocumento.RegistroDoc.Add(regdoc);
                            }
                        }

                     
                        settings = new XmlWriterSettings();
                        settings.Indent = true;
                        settings.Encoding = Encoding.UTF8;
                        settings.OmitXmlDeclaration = true;                        
                        writer = XmlWriter.Create(file.FullName, settings);

                        XmlSerializerNamespaces ns = new XmlSerializerNamespaces();                        
                        ns.Add("", @"http://localhost/elegibilidad");
                        
                        serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101));

                        serializer.Serialize(writer, m101, ns);
                        writer.Close();

                        mensaje_error = ValidateSchema(file.FullName, System.Environment.CurrentDirectory + param_cnmc.GetValue("xsd_m101"));
                        

                        if (mensaje_error != "")
                        {
                           // lista_log.Add(System.Environment.NewLine);
                            lista_log.Add("Hoja: " + p.hoja + " CUPS: " + p.cups22 +
                                        " --> " + mensaje_error);
                            // delete---
                        }
                        else
                        {
                            //lista_log.Add(System.Environment.NewLine);
                            lista_log.Add("Hoja: " + p.hoja + " CUPS: " + p.cups22 +
                                        " --> " + "XML generado correctamente");
                        }

                    }

                    if (p.hoja == "BAJA")
                    {
                        //NO PROCESAMOS ESTA PESTAÑA HASTA FINALIZAR IMPLEMENTACION
                        //FORZAMOS BREAK
                       // lista_log.Add("La pestaña BAJA  no se procesará, está deshabilitada temporalmente");
                        // continue;

                        file = new FileInfo(param.GetValue("ruta_salida_xml") + "B101"
                            + "_" + p.cups22 + "_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xml");
                        

                        EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeB101 b101 =
                            new EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeB101();
                        #region Cabecera

                        EndesaEntity.cnmc.V21_2019_12_17.Cabecera cabecera = new EndesaEntity.cnmc.V21_2019_12_17.Cabecera();

                        cabecera.CodigoREEEmpresaEmisora = p.empresa_emisora;

                        cabecera.CodigoREEEmpresaDestino = p.distribuidora;
                        cabecera.CodigoDelProceso = "B1";
                        cabecera.CodigoDePaso = "01";
                        cabecera.CodigoDeSolicitud = DateTime.Now.ToString("yyMMddHHmmss");
                        cabecera.SecuencialDeSolicitud = "00";
                        cabecera.FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
                        cabecera.CUPS = p.cups22;
                        b101.Cabecera = cabecera;
                        #endregion

                        #region DatosSolicitud

                        b101.BajaSuspension.DatosSolicitud.IndActivacion = p.ind_activacion;

                        if (p.ind_activacion == "F")
                            b101.BajaSuspension.DatosSolicitud.fechaPrevistaAccion = 
                                p.fecha_activacion.ToString("yyyy-MM-dd");

                        b101.BajaSuspension.DatosSolicitud.Motivo = p.motivo_baja;

                        #endregion

                        #region Cliente
                        EndesaEntity.cnmc.V30_2022_21_01.Cliente cliente = new EndesaEntity.cnmc.V30_2022_21_01.Cliente();

                        cliente.IdCliente.TipoIdentificador = p.tipo_identificador;
                        //cliente.IdCliente.TipoIdentificador = cnmc.GetTipo_Identificador(p.tipo_identificador);
                        cliente.IdCliente.Identificador = p.n_identificador;

                        // Leer primer carácter
                        cliente.IdCliente.TipoPersona = p.tipo_persona;

                        b101.BajaSuspension.Cliente = cliente;
                        #endregion

                        #region Nombre
                        if (p.tipo_persona == "F")
                        {
                            b101.BajaSuspension.Cliente.Nombre.NombreDePila = p.nombre_de_pila;
                            b101.BajaSuspension.Cliente.Nombre.PrimerApellido = p.primer_apellido;
                        }
                        else
                        {
                            b101.BajaSuspension.Cliente.Nombre.RazonSocial = p.razon_social;
                        }
                        #endregion

                        b101.BajaSuspension.Cliente.Telefono.PrefijoPais = "0034";
                        b101.BajaSuspension.Cliente.Telefono.Numero = p.telefono;

                        b101.BajaSuspension.Contacto.PersonaDeContacto = p.persona_contacto;
                        
                        b101.BajaSuspension.Contacto.Telefono.PrefijoPais = "0034";
                        b101.BajaSuspension.Contacto.Telefono.Numero = p.tlf_contacto;
                                                                        
                        // TODO:
                        foreach (Documentacion pp in p.lista_documentacion)
                        {
                            RegistroDoc regdoc = new RegistroDoc();
                            regdoc.TipoDocAportado = pp.tipo_doc_aportado;
                            regdoc.DireccionUrl = pp.direccion_url;
                            b101.BajaSuspension.RegistroDocumento.RegistroDoc.Add(regdoc);
                        }

                        if (!System.IO.Directory.Exists(param.GetValue("ruta_salida_xml")))
                        {
                            System.IO.Directory.CreateDirectory(param.GetValue("ruta_salida_xml"));
                        }


                        settings = new XmlWriterSettings();
                        settings.Indent = true;
                        settings.Encoding = Encoding.UTF8;
                        settings.OmitXmlDeclaration = true;
                        writer = XmlWriter.Create(file.FullName, settings);

                        XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                        ns.Add("", @"http://localhost/elegibilidad");
                        serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeB101));

                        serializer.Serialize(writer, b101, ns);
                        writer.Close();

                        mensaje_error = ValidateSchema(file.FullName, 
                            System.Environment.CurrentDirectory + param_cnmc.GetValue("xsd_b101"));

                        if (mensaje_error != "")
                        {
                           // lista_log.Add(System.Environment.NewLine);
                            lista_log.Add("Hoja: " + p.hoja + " CUPS: " + p.cups22 +
                                        " --> " + mensaje_error);
                           // file.Delete();
                        }
                        else
                        {
                            //lista_log.Add(System.Environment.NewLine);
                            lista_log.Add("Hoja: " + p.hoja + " CUPS: " + p.cups22 +
                                        " --> " + "XML generado correctamente");
                        }

                    }

                }

            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message,
                       "GeneraXML",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Error);
            }
        }

        public string ValidateSchema(string xmlPath, string xsdPath)
        {

            //string mensaje = "";
            string mensaje = string.Empty;
            XmlDocument xml = new XmlDocument();
            xml.Load(xmlPath);

            xml.Schemas.Add(null, xsdPath);

            //try
            //{
            //    xml.Validate(null);
            //}
            //catch (XmlSchemaValidationException e)
             xml.Validate((sender, args) =>
             {
                 //[06/03/2025 GUS]: if(!e.Message.Contains("http://localhost/elegibilidad"))
                 //if (e.Message.Contains("http://localhost/elegibilidad"))
                 //        return e.Message;
                 //}
                 //return mensaje;
                 mensaje += args.Severity + ": " + args.Message + Environment.NewLine;
             });

            return mensaje.Trim();
        }

        private void BorradoDatosMySQL()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "delete from extrasistemas_excel_ad where created_by = '" + System.Environment.UserName + "'";
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }



        public List<EndesaEntity.xml.Plantilla_Extrasistemas_total_registros> Totales_Registros()
        {

            List<EndesaEntity.xml.Plantilla_Extrasistemas_total_registros> lista =
                new List<EndesaEntity.xml.Plantilla_Extrasistemas_total_registros>();
            foreach (KeyValuePair<string, EndesaEntity.xml.Plantilla_Extrasistemas_total_registros> p in dic_total_registros)
            {
                lista.Add(p.Value);
            }
            return lista;
        }

        private Dictionary<string, EndesaEntity.xml.Plantilla_Extrasistemas_total_registros> Inicializa_Lista_Totales()
        {
            Dictionary<string, EndesaEntity.xml.Plantilla_Extrasistemas_total_registros> d =
                new Dictionary<string, EndesaEntity.xml.Plantilla_Extrasistemas_total_registros>();


            EndesaEntity.xml.Plantilla_Extrasistemas_total_registros c = 
                new EndesaEntity.xml.Plantilla_Extrasistemas_total_registros();

            c.hoja = "AD";
            c.descripcion = "Alta Directa A301";
            c.registros = 0;
            d.Add(c.hoja,c);

            c = new EndesaEntity.xml.Plantilla_Extrasistemas_total_registros();
            c.hoja = "ACC";
            c.descripcion = "ACC C101";
            c.registros = 0;
            d.Add(c.hoja, c);

            c = new EndesaEntity.xml.Plantilla_Extrasistemas_total_registros();
            c.hoja = "ACC + CAMBIOS";
            c.descripcion = "ACC+ C201";
            c.registros = 0;
            d.Add(c.hoja, c);

            c = new EndesaEntity.xml.Plantilla_Extrasistemas_total_registros();
            c.hoja = "MOD";
            c.descripcion = "Modificación M101";
            c.registros = 0;
            d.Add(c.hoja, c);

            c = new EndesaEntity.xml.Plantilla_Extrasistemas_total_registros();
            c.hoja = "BAJA";
            c.descripcion = "Baja B101";
            c.registros = 0;
            d.Add(c.hoja, c);

            return d;

        }



    }
}
