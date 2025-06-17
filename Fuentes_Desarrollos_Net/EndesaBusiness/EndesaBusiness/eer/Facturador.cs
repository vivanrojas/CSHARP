using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace EndesaBusiness.eer
{
    public class Facturador
    {
        EndesaBusiness.calendarios.UtilidadesCalendario utilFecha;
        EndesaBusiness.medida.CurvasEER curvas;
        EndesaBusiness.utilidades.Param param;
        
        EndesaBusiness.eer.Coef_Excesos_Potencia coef_excesos;
        EndesaBusiness.eer.Precios_Excesos_Potencia precio_excesos;

        EndesaBusiness.ree.SuplementoTerritorial suplementoTerritorial;
        EndesaBusiness.facturacion.CosPhi cosphi;
        EndesaBusiness.eer.RegularizacionPrecios regPrecios;
        bool firstOnly = true;
        

        double[] vectorAci;
        public Facturador()
        {
            param = new utilidades.Param("eer_param", MySQLDB.Esquemas.CON);
        }

        public Facturador(DateTime fd, DateTime fh, List<EndesaEntity.punto_suministro.PuntoSuministro> lista_PS, 
            EndesaBusiness.eer.Inventario inventario, bool exportar_calculos)
        {
            utilidades.Fechas utilFechas = new utilidades.Fechas();

            string fechaHora_gen_factura = "";
            EndesaBusiness.calendarios.Calendario cal;
            EndesaEntity.eer.Factura factura;
            EndesaEntity.eer.FacturaDetalle facturaDetalle;
            double meses = 0;
            int dias = 0;
            int dias_anio = 0;

            EndesaBusiness.eer.FacturasEER factEER;

            List<string> lista_cups22 = new List<string>();
            int numLineaFactura;
            double[] vectorPotenciasMaximasRegistradas;
            double[] vectorExcesos;

            double importe_termino_energia = 0;
            double importe_potencias = 0;
            double importe_excesos_potencia = 0;
            double importe_energia_reactiva = 0;
            double total_cargos = 0;
            double cargos_potencia = 0;
            double cargos_activa = 0;
            
            EndesaEntity.DatosPeaje peaje;

            double total_ATR_Cargo_Energia = 0;
            double total_ATR_Cargo_Potencia = 0;
            double total_TP = 0;
            double total_TE = 0;


            EndesaBusiness.eer.PreciosEnergia precios_energia = new PreciosEnergia(fd, fh);

            // Suplemento Territorial
            suplementoTerritorial = new ree.SuplementoTerritorial();

            //EndesaBusiness.ree.PreciosRegulados preciosRegulados = new ree.PreciosRegulados(fd, fh);
            EndesaBusiness.medida.DatosPeajesEER peajesEER = new medida.DatosPeajesEER(fd, fh);

            coef_excesos = new Coef_Excesos_Potencia(fd, fh);
            precio_excesos = new Precios_Excesos_Potencia(fd, fh);


            regPrecios = new RegularizacionPrecios(Convert.ToInt32(fd.ToString("yyyyMM")));

            // Excesos de reactiva
            cosphi = new facturacion.CosPhi(fd, fh);
            double new_totalPotenciaPeriodos = 0;

            // Descuento RDL17/2021
            double descuento = 0;
            DateTime fecha_descuento = new DateTime(2021, 09, 15);
            EndesaBusiness.ree.PreciosRegulados preciosRegulados_descuento = new ree.PreciosRegulados();
            double total_ATR_Cargo_Energia_descuento = 0;
            double total_ATR_Cargo_Potencia_descuento = 0;
            double total_TP_descuento = 0;
            double total_TE_descuento = 0;

            EndesaEntity.ree.PreciosRegulados precio_regulado_ATR_Cargo_Energia_descuento = new EndesaEntity.ree.PreciosRegulados();
            EndesaEntity.ree.PreciosRegulados precio_regulado_TE_descuento = new EndesaEntity.ree.PreciosRegulados();

            EndesaEntity.ree.PreciosRegulados atr_cargo_potencia_descuento = new EndesaEntity.ree.PreciosRegulados();


            double new_totalPotenciaPeriodos_descuento = 0;
            // FIN Descuento RDL17/2021



            param = new utilidades.Param("eer_param", MySQLDB.Esquemas.CON);

            for (int i = 0; i < lista_PS.Count; i++)
            {
                importe_termino_energia = 0;
                importe_potencias = 0;
                importe_excesos_potencia = 0;
                importe_energia_reactiva = 0;
                total_cargos = 0;
                cargos_potencia = 0;
                cargos_activa = 0;
                               

                total_ATR_Cargo_Energia = 0;
                total_ATR_Cargo_Potencia = 0;
                total_TP = 0;
                total_TE = 0;

                total_ATR_Cargo_Energia_descuento = 0;
                total_ATR_Cargo_Potencia_descuento = 0;
                total_TP_descuento = 0;
                total_TE_descuento = 0;
                new_totalPotenciaPeriodos_descuento = 0;



                EndesaBusiness.ree.PreciosRegulados preciosRegulados = new ree.PreciosRegulados(lista_PS[i].fecha_inicio, lista_PS[i].fecha_fin);
                //if(param.GetValue("CD_RDL17/2021") == "S")
                if(fd < new DateTime(2022,01,01))
                     preciosRegulados_descuento = new ree.PreciosRegulados(fecha_descuento, fecha_descuento);


                // Para cuando el periodo de consumo no es un mes natural
                meses = CoeficienteMes(lista_PS[i].fecha_inicio, lista_PS[i].fecha_fin);
                //dias = Convert.ToInt32((fh - fd).TotalDays + 1);
                dias = Convert.ToInt32((lista_PS[i].fecha_fin - lista_PS[i].fecha_inicio).TotalDays + 1);
                dias_anio = utilFechas.EsAnioBisiesto(fd) ? 366 : 365;

                numLineaFactura = 0;
                cal = new calendarios.Calendario(lista_PS[i].fecha_inicio, lista_PS[i].fecha_fin);
                for (int l = 0; l < lista_PS[i].lista_puntos_medida_principales.Count; l++)
                    lista_cups22.Add(lista_PS[i].lista_puntos_medida_principales[l].cups22);


                lista_PS[i].precios_energia = precios_energia.GetPrecio(lista_PS[i].cups20, lista_PS[i].fecha_inicio, lista_PS[i].fecha_fin);
                if(lista_PS[i].precios_energia == null)
                {
                    MessageBox.Show("No se han encontrado precios para el CUPS: " + lista_PS[i].cups20
                        + System.Environment.NewLine
                        + "La facturación no puede continuar.",
                     "Sin precios",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Error);
                    break;
                }

                curvas = new medida.CurvasEER(lista_PS[i], lista_PS[i].fecha_inicio, lista_PS[i].fecha_fin, "R", true, false);
                if (curvas.curvaCompleta || peajesEER.HayDatos(lista_PS[i].cups20))
                {
                    peaje = peajesEER.GetDatosPeaje(lista_PS[i].cups20);
                    factEER = new FacturasEER();

                    EndesaEntity.Inventario cliente = inventario.GetCliente(lista_PS[i].cups20, lista_PS[i].fecha_inicio, lista_PS[i].fecha_fin);
                    if (cliente != null)
                    {
                        factura = new EndesaEntity.eer.Factura();
                        factura.nombre_cliente = cliente.nombre_cliente;
                        factura.nif = cliente.nif;
                        factura.cupsree = lista_PS[i].cups20;
                        factura.direccion_facturacion = cliente.direccion_facturacion;
                        factura.direccion_suministro = lista_PS[i].direccion.direccion_completa;

                        factura.id_factura = UltimaFactura() + 1;
                        factura.fecha_consumo_desde = lista_PS[i].fecha_inicio;
                        factura.fecha_consumo_hasta = lista_PS[i].fecha_fin;

                        factura.consumo_activa = curvas.totalEnergiaActiva;
                        factura.consumo_reactiva = curvas.totalEnergiaReactiva;
                        factura.tarifa = lista_PS[i].tarifa.tarifa;


                        lista_PS[i].dic_cc = curvas.GetCurva(lista_PS[i].cups20);

                        if (curvas.curvaCompleta)
                            cal.CalculaDatosMedida(lista_PS[i],
                                curvas.curvaCuartoHorariaActiva,
                                curvas.curvaCuartoHorariaReactiva,
                                curvas.curvaCuartoHorariaPotencias,
                                curvas.curvaCuartoHorariaDias);
                        else
                        {
                            peaje = peajesEER.GetDatosPeaje(lista_PS[i].cups20);
                            for (int q = 1; q <= 6; q++)
                            {
                                factura.consumo_activa += peaje.activa[q];
                                factura.consumo_reactiva += peaje.reactiva[q];
                            }
                                
                        }


                        #region Termino Energia Variable
                        firstOnly = true;

                        EndesaEntity.ree.PreciosRegulados precio_regulado_ATR_Cargo_Energia
                            = preciosRegulados.GetPrecioRegulado("ATR_Cargo_Energia", lista_PS[i].tarifa.tarifa);

                        EndesaEntity.ree.PreciosRegulados precio_regulado_TE
                            = preciosRegulados.GetPrecioRegulado("TE", lista_PS[i].tarifa.tarifa);

                        //if (param.GetValue("CD_RDL17/2021") == "S")
                        if (fd < new DateTime(2022, 01, 01))
                        {
                            precio_regulado_ATR_Cargo_Energia_descuento
                           = preciosRegulados_descuento.GetPrecioRegulado("ATR_Cargo_Energia", lista_PS[i].tarifa.tarifa);

                            precio_regulado_TE_descuento
                                = preciosRegulados_descuento.GetPrecioRegulado("TE", lista_PS[i].tarifa.tarifa);
                        }


                        if (curvas.curvaCompleta)
                        {
                            for (int pt = 1; pt < cal.energiaActivaPorPeriodo.Count(); pt++)
                            {
                                if (cal.energiaActivaPorPeriodo[pt] != 0)
                                {
                                    numLineaFactura++;
                                    facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                    facturaDetalle.linea_factura = numLineaFactura;
                                    facturaDetalle.producto = "L01";
                                    facturaDetalle.concepto = "Término de Energía Variable";
                                    facturaDetalle.cantidad = cal.energiaActivaPorPeriodo[pt];
                                    facturaDetalle.unidad_cantidad = "KWh";
                                    facturaDetalle.precio = lista_PS[i].precios_energia.precios_periodo[pt];
                                    facturaDetalle.unidad_precio = "Eur/KWh";
                                    facturaDetalle.total = cal.energiaActivaPorPeriodo[pt] * lista_PS[i].precios_energia.precios_periodo[pt];
                                    facturaDetalle.descripcion = "P" + pt + ": "
                                        + string.Format("{0:#,##0}", facturaDetalle.cantidad)
                                        + " kWh x "
                                        + facturaDetalle.precio
                                        + " Eur/kWh = "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.total)
                                        + " Eur";
                                    factura.lista_factura_detalle.Add(facturaDetalle);
                                    importe_termino_energia = importe_termino_energia + facturaDetalle.total;
                                    factura.base_imponible = factura.base_imponible + facturaDetalle.total;
                                    factura.base_imponible_ie = factura.base_imponible_ie + facturaDetalle.total;

                                }

                                if (cal.energiaActivaPorPeriodo[pt] != 0)
                                {
                                    total_ATR_Cargo_Energia =
                                       total_ATR_Cargo_Energia + (cal.energiaActivaPorPeriodo[pt] * precio_regulado_ATR_Cargo_Energia.periodo_tarifario[pt]);

                                    total_TE =
                                        total_TE + (cal.energiaActivaPorPeriodo[pt] * precio_regulado_TE.periodo_tarifario[pt]);

                                    //if (param.GetValue("CD_RDL17/2021") == "S")
                                    if (fd < new DateTime(2022, 01, 01))
                                    {
                                        total_ATR_Cargo_Energia_descuento =
                                      total_ATR_Cargo_Energia_descuento + 
                                      (cal.energiaActivaPorPeriodo[pt] * precio_regulado_ATR_Cargo_Energia_descuento.periodo_tarifario[pt]);

                                        total_TE_descuento =
                                            total_TE_descuento + 
                                            (cal.energiaActivaPorPeriodo[pt] * precio_regulado_TE_descuento.periodo_tarifario[pt]);
                                    }

                                }


                            }


                        }
                        else
                        {


                            for (int pt = 1; pt <= lista_PS[i].tarifa.numPeriodosTarifarios; pt++)
                            {
                                if (peaje.activa[pt] != 0)
                                {
                                    numLineaFactura++;
                                    facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                    facturaDetalle.linea_factura = numLineaFactura;
                                    facturaDetalle.producto = "L01";
                                    facturaDetalle.concepto = "Término de Energía Variable";
                                    facturaDetalle.cantidad = peaje.activa[pt];
                                    facturaDetalle.unidad_cantidad = "KWh";
                                    facturaDetalle.precio = lista_PS[i].precios_energia.precios_periodo[pt];
                                    facturaDetalle.unidad_precio = "Eur/KWh";
                                    facturaDetalle.total = peaje.activa[pt] * lista_PS[i].precios_energia.precios_periodo[pt];
                                    facturaDetalle.descripcion = "P" + pt + ": "
                                        + string.Format("{0:#,##0}", facturaDetalle.cantidad)
                                        + " kWh x "
                                        + facturaDetalle.precio
                                        + " Eur/kWh = "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.total)
                                        + " Eur";
                                    factura.lista_factura_detalle.Add(facturaDetalle);
                                    importe_termino_energia = importe_termino_energia + facturaDetalle.total;
                                    factura.base_imponible = factura.base_imponible + facturaDetalle.total;
                                    factura.base_imponible_ie = factura.base_imponible_ie + facturaDetalle.total;
                                    factura.consumos_periodos_activa[pt] = peaje.activa[pt];

                                }

                                if (peaje.activa[pt] != 0)
                                {
                                    total_ATR_Cargo_Energia =
                                       total_ATR_Cargo_Energia + (peaje.activa[pt] * precio_regulado_ATR_Cargo_Energia.periodo_tarifario[pt]);

                                    total_TE =
                                        total_TE + (peaje.activa[pt] * precio_regulado_TE.periodo_tarifario[pt]);

                                    //if (param.GetValue("CD_RDL17/2021") == "S")
                                    if(fd < new DateTime(2022,01,01))
                                    {
                                        total_ATR_Cargo_Energia_descuento =
                                      total_ATR_Cargo_Energia_descuento +
                                      (peaje.activa[pt] * precio_regulado_ATR_Cargo_Energia_descuento.periodo_tarifario[pt]);

                                        total_TE_descuento =
                                            total_TE_descuento +
                                            (peaje.activa[pt] * precio_regulado_TE_descuento.periodo_tarifario[pt]);

                                       
                                    }

                                }

                            }
                        }
                        #endregion

                        #region EnergiaReactiva


                        double excesosReactiva = 0;

                        if (curvas.curvaCompleta)
                        {
                            excesosReactiva = cosphi.ImporteCosPhi(lista_PS[i].tarifa.numPeriodosTarifarios,
                                cal.energiaActivaPorPeriodo, cal.energiaReactivaPorPeriodo);

                            if (excesosReactiva > 0)
                            {

                                numLineaFactura++;
                                facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                facturaDetalle.linea_factura = numLineaFactura;
                                facturaDetalle.producto = "L44";
                                facturaDetalle.concepto = "Energía Reactiva";
                                facturaDetalle.total = excesosReactiva;
                                factura.lista_factura_detalle.Add(facturaDetalle);

                                factura.base_imponible = factura.base_imponible + facturaDetalle.total;
                                factura.base_imponible_ie = factura.base_imponible_ie + facturaDetalle.total;

                                importe_energia_reactiva = facturaDetalle.total;
                            }
                        }
                        else
                        {
                            if (peaje.importe_excesos_reactiva > 0)
                            {

                                numLineaFactura++;
                                facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                facturaDetalle.linea_factura = numLineaFactura;
                                facturaDetalle.producto = "L44";
                                facturaDetalle.concepto = "Energía Reactiva";
                                facturaDetalle.total = peaje.importe_excesos_reactiva;                               

                                factura.lista_factura_detalle.Add(facturaDetalle);

                                factura.base_imponible = factura.base_imponible + facturaDetalle.total;
                                factura.base_imponible_ie = factura.base_imponible_ie + facturaDetalle.total;

                                importe_energia_reactiva = facturaDetalle.total;

                            }
                        }




                        #endregion


                        #region Potencia Periodos
                        if (fd < new DateTime(2021, 06, 01))
                        {
                            EndesaEntity.ree.PreciosRegulados precioPotencias
                            = preciosRegulados.GetPrecioRegulado("TerminosPotencia", lista_PS[i].tarifa.tarifa);

                            double totalPotenciaPeriodos = 0;
                            if (curvas.curvaCompleta)
                            {
                                for (int pt = 1; pt <= lista_PS[i].tarifa.numPeriodosTarifarios; pt++)
                                {
                                    numLineaFactura++;
                                    facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                    facturaDetalle.linea_factura = numLineaFactura;
                                    facturaDetalle.producto = "L34";
                                    facturaDetalle.concepto = "Facturación Potencia Periodos";
                                    facturaDetalle.unidad_cantidad = "KW";
                                    facturaDetalle.unidad_precio = "Eur/KW";

                                    if (lista_PS[i].tarifa.numPeriodosTarifarios > 3)
                                    {
                                        facturaDetalle.cantidad = lista_PS[i].potecias_contratadas[pt];
                                        facturaDetalle.precio = precioPotencias.periodo_tarifario[pt];
                                        facturaDetalle.total = lista_PS[i].potecias_contratadas[pt] * precioPotencias.periodo_tarifario[pt];
                                    }
                                    else
                                    {
                                        facturaDetalle.cantidad = cal.potenciasaFacturar[pt];
                                        facturaDetalle.precio = precioPotencias.periodo_tarifario[pt];
                                        facturaDetalle.total = cal.potenciasaFacturar[pt] * precioPotencias.periodo_tarifario[pt];
                                    }

                                    facturaDetalle.descripcion = "P" + pt + ": "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.cantidad)
                                        + " kW x "
                                        + facturaDetalle.precio
                                        + " Eur/kW = "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.total)
                                        + " Eur";
                                    factura.lista_factura_detalle.Add(facturaDetalle);
                                    totalPotenciaPeriodos = totalPotenciaPeriodos + (meses * facturaDetalle.total);


                                }
                                importe_potencias = (meses * totalPotenciaPeriodos) / 12;
                                factura.base_imponible = factura.base_imponible + ((totalPotenciaPeriodos) / 12);
                                factura.base_imponible_ie = factura.base_imponible_ie + ((totalPotenciaPeriodos) / 12);
                            }
                            else
                            {
                                numLineaFactura++;
                                facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                facturaDetalle.linea_factura = numLineaFactura;
                                facturaDetalle.producto = "L34";
                                facturaDetalle.concepto = "Facturación Potencia Periodos";
                                facturaDetalle.unidad_cantidad = "KW";
                                facturaDetalle.unidad_precio = "Eur/KW";
                                facturaDetalle.cantidad = 1;
                                facturaDetalle.precio = 1;
                                facturaDetalle.total = peaje.importe_termino_potencia;
                                factura.lista_factura_detalle.Add(facturaDetalle);
                                totalPotenciaPeriodos = totalPotenciaPeriodos + facturaDetalle.total;

                                importe_potencias = peaje.importe_termino_potencia;
                                factura.base_imponible = factura.base_imponible + (peaje.importe_termino_potencia);
                                factura.base_imponible_ie = factura.base_imponible_ie + (peaje.importe_termino_potencia);

                            }
                        }


                        #endregion

                        #region Potencia Periodos 20210601
                        if (fd >= new DateTime(2021, 06, 01))
                        {
                            EndesaEntity.ree.PreciosRegulados new_precioPotencias
                            = preciosRegulados.GetPrecioRegulado("TP", "ATR_Cargo_Potencia", lista_PS[i].tarifa.tarifa);

                            EndesaEntity.ree.PreciosRegulados pptt
                               = preciosRegulados.GetPrecioRegulado("TP", lista_PS[i].tarifa.tarifa);

                            EndesaEntity.ree.PreciosRegulados atr_cargo_potencia
                            = preciosRegulados.GetPrecioRegulado("ATR_Cargo_Potencia", lista_PS[i].tarifa.tarifa);

                            //if (param.GetValue("CD_RDL17/2021") == "S")
                            if (fd < new DateTime(2022, 01, 01))
                            {
                                atr_cargo_potencia_descuento
                                    = preciosRegulados_descuento.GetPrecioRegulado("ATR_Cargo_Potencia", lista_PS[i].tarifa.tarifa);
                                                               
                            }
                            
                            if (curvas.curvaCompleta)
                            {
                                for (int pt = 1; pt <= lista_PS[i].tarifa.numPeriodosTarifarios; pt++)
                                {
                                    numLineaFactura++;
                                    facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                    facturaDetalle.linea_factura = numLineaFactura;
                                    facturaDetalle.producto = "L85";
                                    facturaDetalle.concepto = "Facturación Potencia Periodos";
                                    facturaDetalle.unidad_cantidad = "KW";
                                    facturaDetalle.unidad_precio = "Eur/KW";
                                    if (lista_PS[i].tarifa.numPeriodosTarifarios > 3)
                                    {
                                        facturaDetalle.cantidad = lista_PS[i].potecias_contratadas[pt];
                                        facturaDetalle.precio = new_precioPotencias.periodo_tarifario[pt];
                                        facturaDetalle.total = lista_PS[i].potecias_contratadas[pt] * new_precioPotencias.periodo_tarifario[pt];
                                    }
                                    else
                                    {
                                        facturaDetalle.cantidad = cal.potenciasaFacturar[pt];
                                        facturaDetalle.precio = new_precioPotencias.periodo_tarifario[pt];
                                        facturaDetalle.total = cal.potenciasaFacturar[pt] * new_precioPotencias.periodo_tarifario[pt];
                                    }

                                    facturaDetalle.descripcion = "P" + pt + ": "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.cantidad)
                                        + " kW x "
                                        + facturaDetalle.precio
                                        + " €/Año / kW = "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.total)
                                        + " €/Año";
                                    factura.lista_factura_detalle.Add(facturaDetalle);
                                    new_totalPotenciaPeriodos = new_totalPotenciaPeriodos + ((facturaDetalle.total / dias_anio) * dias);
                                }

                                importe_potencias = new_totalPotenciaPeriodos;
                                factura.base_imponible = factura.base_imponible + new_totalPotenciaPeriodos;
                                factura.base_imponible_ie = factura.base_imponible_ie + new_totalPotenciaPeriodos;
                                factura.facturacion_potencia = new_totalPotenciaPeriodos;


                                #region Nuevos Terminos
                                // Guardamos los importes por separado                              

                                new_totalPotenciaPeriodos = 0;
                                new_totalPotenciaPeriodos_descuento = 0;

                                for (int pt = 1; pt <= lista_PS[i].tarifa.numPeriodosTarifarios; pt++)
                                {
                                    numLineaFactura++;
                                    facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                    facturaDetalle.linea_factura = numLineaFactura;
                                    facturaDetalle.producto = "TP";
                                    facturaDetalle.concepto = "TP";
                                    facturaDetalle.unidad_cantidad = "KW";
                                    facturaDetalle.unidad_precio = "Eur/KW";

                                    if (lista_PS[i].tarifa.numPeriodosTarifarios > 3)
                                    {
                                        facturaDetalle.cantidad = lista_PS[i].potecias_contratadas[pt];
                                        facturaDetalle.precio = pptt.periodo_tarifario[pt];
                                        facturaDetalle.total = lista_PS[i].potecias_contratadas[pt] * pptt.periodo_tarifario[pt];

                                    }
                                    else
                                    {
                                        facturaDetalle.cantidad = cal.potenciasaFacturar[pt];
                                        facturaDetalle.precio = pptt.periodo_tarifario[pt];
                                        facturaDetalle.total = cal.potenciasaFacturar[pt] * pptt.periodo_tarifario[pt];
                                    }

                                    new_totalPotenciaPeriodos = new_totalPotenciaPeriodos + facturaDetalle.total;

                                    facturaDetalle.descripcion = "P" + pt + ": "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.cantidad)
                                        + " kW x "
                                        + facturaDetalle.precio
                                        + " Eur/kW = "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.total)
                                        + " Eur";
                                    factura.lista_factura_detalle.Add(facturaDetalle);

                                }

                                total_TP = (new_totalPotenciaPeriodos / dias_anio) * dias;

                                new_totalPotenciaPeriodos = 0;

                                for (int pt = 1; pt <= lista_PS[i].tarifa.numPeriodosTarifarios; pt++)
                                {
                                    numLineaFactura++;
                                    facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                    facturaDetalle.linea_factura = numLineaFactura;
                                    facturaDetalle.producto = "ATR_Cargo_Potencia";
                                    facturaDetalle.concepto = "ATR_Cargo_Potencia";
                                    facturaDetalle.unidad_cantidad = "KW";
                                    facturaDetalle.unidad_precio = "Eur/KW";
                                    if (lista_PS[i].tarifa.numPeriodosTarifarios > 3)
                                    {
                                        facturaDetalle.cantidad = lista_PS[i].potecias_contratadas[pt];
                                        facturaDetalle.precio = atr_cargo_potencia.periodo_tarifario[pt];
                                        facturaDetalle.total =
                                            ((lista_PS[i].potecias_contratadas[pt] * atr_cargo_potencia.periodo_tarifario[pt]) / dias_anio) * dias;
                                    }
                                    else
                                    {
                                        facturaDetalle.cantidad = cal.potenciasaFacturar[pt];
                                        facturaDetalle.precio = atr_cargo_potencia.periodo_tarifario[pt];
                                        facturaDetalle.total = cal.potenciasaFacturar[pt] * atr_cargo_potencia.periodo_tarifario[pt];
                                    }

                                    total_ATR_Cargo_Potencia =
                                        total_ATR_Cargo_Potencia + facturaDetalle.total;

                                    facturaDetalle.descripcion = "P" + pt + ": "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.cantidad)
                                        + " kW x "
                                        + facturaDetalle.precio
                                        + " Eur/kW = "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.total)
                                        + " Eur";
                                    factura.lista_factura_detalle.Add(facturaDetalle);

                                }

                                //if (param.GetValue("CD_RDL17/2021") == "S")
                                if (fd < new DateTime(2022, 01, 01))
                                {
                                    for (int pt = 1; pt <= lista_PS[i].tarifa.numPeriodosTarifarios; pt++)
                                    {
                                        //numLineaFactura++;
                                        facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                        facturaDetalle.linea_factura = numLineaFactura;
                                        facturaDetalle.producto = "ATR_Cargo_Potencia";
                                        facturaDetalle.concepto = "ATR_Cargo_Potencia";
                                        facturaDetalle.unidad_cantidad = "KW";
                                        facturaDetalle.unidad_precio = "Eur/KW";
                                        if (lista_PS[i].tarifa.numPeriodosTarifarios > 3)
                                        {
                                            facturaDetalle.cantidad = lista_PS[i].potecias_contratadas[pt];
                                            facturaDetalle.precio = atr_cargo_potencia_descuento.periodo_tarifario[pt];
                                            facturaDetalle.total =
                                                ((lista_PS[i].potecias_contratadas[pt] * atr_cargo_potencia_descuento.periodo_tarifario[pt]) / dias_anio) * dias;
                                        }
                                        else
                                        {
                                            facturaDetalle.cantidad = cal.potenciasaFacturar[pt];
                                            facturaDetalle.precio = atr_cargo_potencia_descuento.periodo_tarifario[pt];
                                            facturaDetalle.total = cal.potenciasaFacturar[pt] * atr_cargo_potencia_descuento.periodo_tarifario[pt];
                                        }

                                        total_ATR_Cargo_Potencia_descuento =
                                            total_ATR_Cargo_Potencia_descuento + facturaDetalle.total;

                                        //facturaDetalle.descripcion = "P" + pt + ": "
                                        //    + string.Format("{0:#,##0.00}", facturaDetalle.cantidad)
                                        //    + " kW x "
                                        //    + facturaDetalle.precio
                                        //    + " Eur/kW = "
                                        //    + string.Format("{0:#,##0.00}", facturaDetalle.total)
                                        //    + " Eur";
                                        //factura.lista_factura_detalle.Add(facturaDetalle);

                                    }
                                }


                                #endregion

                            }
                            else
                            {
                                new_totalPotenciaPeriodos = 0;

                                for (int pt = 1; pt <= lista_PS[i].tarifa.numPeriodosTarifarios; pt++)
                                {
                                    numLineaFactura++;
                                    facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                    facturaDetalle.linea_factura = numLineaFactura;
                                    facturaDetalle.producto = "L85";
                                    facturaDetalle.concepto = "Facturación Potencia Periodos";
                                    facturaDetalle.unidad_cantidad = "KW";
                                    facturaDetalle.unidad_precio = "Eur/KW";
                                    facturaDetalle.cantidad = 1;
                                    facturaDetalle.precio = 1;

                                    if (lista_PS[i].tarifa.numPeriodosTarifarios > 3)
                                    {
                                        facturaDetalle.cantidad = lista_PS[i].potecias_contratadas[pt];
                                        facturaDetalle.precio = new_precioPotencias.periodo_tarifario[pt];
                                        facturaDetalle.total = lista_PS[i].potecias_contratadas[pt] * new_precioPotencias.periodo_tarifario[pt];
                                    }
                                    else
                                    {
                                        facturaDetalle.cantidad = cal.potenciasaFacturar[pt];
                                        facturaDetalle.precio = new_precioPotencias.periodo_tarifario[pt];
                                        facturaDetalle.total = cal.potenciasaFacturar[pt] * new_precioPotencias.periodo_tarifario[pt];
                                    }

                                    

                                    facturaDetalle.descripcion = "P" + pt + ": "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.cantidad)
                                        + " kW x "
                                        + facturaDetalle.precio
                                        + " €/Año / kW = "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.total)
                                        + " €/Año";
                                    factura.lista_factura_detalle.Add(facturaDetalle);
                                    new_totalPotenciaPeriodos = new_totalPotenciaPeriodos + ((facturaDetalle.total / dias_anio) * dias);

                                }


                                //importe_potencias = peaje.importe_termino_potencia;
                                //factura.base_imponible = factura.base_imponible + (peaje.importe_termino_potencia);
                                //factura.base_imponible_ie = factura.base_imponible_ie + (peaje.importe_termino_potencia);

                                importe_potencias = new_totalPotenciaPeriodos;
                                factura.base_imponible = factura.base_imponible + new_totalPotenciaPeriodos;
                                factura.base_imponible_ie = factura.base_imponible_ie + new_totalPotenciaPeriodos;

                                // Guardamos los importes por separado                              

                                new_totalPotenciaPeriodos = 0;

                                for (int pt = 1; pt <= lista_PS[i].tarifa.numPeriodosTarifarios; pt++)
                                {
                                    numLineaFactura++;
                                    facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                    facturaDetalle.linea_factura = numLineaFactura;
                                    facturaDetalle.producto = "TP";
                                    facturaDetalle.concepto = "TP";
                                    facturaDetalle.unidad_cantidad = "KW";
                                    facturaDetalle.unidad_precio = "Eur/KW";

                                    if (lista_PS[i].tarifa.numPeriodosTarifarios > 3)
                                    {
                                        facturaDetalle.cantidad = lista_PS[i].potecias_contratadas[pt];
                                        facturaDetalle.precio = pptt.periodo_tarifario[pt];
                                        facturaDetalle.total = lista_PS[i].potecias_contratadas[pt] * pptt.periodo_tarifario[pt];

                                    }
                                    else
                                    {
                                        facturaDetalle.cantidad = cal.potenciasaFacturar[pt];
                                        facturaDetalle.precio = pptt.periodo_tarifario[pt];
                                        facturaDetalle.total = cal.potenciasaFacturar[pt] * pptt.periodo_tarifario[pt];
                                    }

                                    new_totalPotenciaPeriodos = new_totalPotenciaPeriodos + facturaDetalle.total;

                                    facturaDetalle.descripcion = "P" + pt + ": "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.cantidad)
                                        + " kW x "
                                        + facturaDetalle.precio
                                        + " Eur/kW = "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.total)
                                        + " Eur";
                                    factura.lista_factura_detalle.Add(facturaDetalle);

                                }

                                total_TP = (new_totalPotenciaPeriodos / dias_anio) * dias;



                                new_totalPotenciaPeriodos = 0;

                                for (int pt = 1; pt <= lista_PS[i].tarifa.numPeriodosTarifarios; pt++)
                                {
                                    numLineaFactura++;
                                    facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                    facturaDetalle.linea_factura = numLineaFactura;
                                    facturaDetalle.producto = "ATR_Cargo_Potencia";
                                    facturaDetalle.concepto = "ATR_Cargo_Potencia";
                                    facturaDetalle.unidad_cantidad = "KW";
                                    facturaDetalle.unidad_precio = "Eur/KW";
                                    if (lista_PS[i].tarifa.numPeriodosTarifarios > 3)
                                    {
                                        facturaDetalle.cantidad = lista_PS[i].potecias_contratadas[pt];
                                        facturaDetalle.precio = atr_cargo_potencia.periodo_tarifario[pt];
                                        facturaDetalle.total =
                                            ((lista_PS[i].potecias_contratadas[pt] * atr_cargo_potencia.periodo_tarifario[pt]) / dias_anio) * dias;
                                    }
                                    else
                                    {
                                        facturaDetalle.cantidad = cal.potenciasaFacturar[pt];
                                        facturaDetalle.precio = atr_cargo_potencia.periodo_tarifario[pt];
                                        facturaDetalle.total = cal.potenciasaFacturar[pt] * atr_cargo_potencia.periodo_tarifario[pt];
                                    }

                                    total_ATR_Cargo_Potencia =
                                        total_ATR_Cargo_Potencia + facturaDetalle.total;

                                    facturaDetalle.descripcion = "P" + pt + ": "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.cantidad)
                                        + " kW x "
                                        + facturaDetalle.precio
                                        + " Eur/kW = "
                                        + string.Format("{0:#,##0.00}", facturaDetalle.total)
                                        + " Eur";
                                    factura.lista_factura_detalle.Add(facturaDetalle);

                                }

                            }
                        }


                    

                        #endregion

                        #region SSTT
                        if (lista_PS[i].sstt)
                        {
                            numLineaFactura++;
                            facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                            facturaDetalle.linea_factura = numLineaFactura;
                            facturaDetalle.producto = "SSTT";
                            facturaDetalle.concepto = "Suplemento Territorial de Peaje";
                            facturaDetalle.cantidad = 1;
                            facturaDetalle.unidad_cantidad = "un.";
                            facturaDetalle.precio = lista_PS[i].importe_sstt;
                            facturaDetalle.unidad_precio = "Eur";
                            facturaDetalle.total = lista_PS[i].importe_sstt;
                            factura.lista_factura_detalle.Add(facturaDetalle);

                            factura.base_imponible = factura.base_imponible + facturaDetalle.total;
                            factura.base_imponible_ie = factura.base_imponible_ie + facturaDetalle.total;
                        }


                        #endregion

                        #region Excesos

                        vectorExcesos = CargaExcesos(cal, lista_PS[i]);
                        vectorPotenciasMaximasRegistradas = CargaPotenciasMaximasRegistradas(cal, lista_PS[i]);
                        vectorAci = new double[lista_PS[i].tarifa.numPeriodosTarifarios + 1];


                        double parcial = 0;
                        double suma = 0;
                        numLineaFactura++;
                        facturaDetalle = new EndesaEntity.eer.FacturaDetalle();

                        facturaDetalle.linea_factura = numLineaFactura;
                        facturaDetalle.producto = "REPA";
                        facturaDetalle.concepto = "Recargo por Excesos de Potencia";
                        facturaDetalle.cantidad = 1;
                        facturaDetalle.unidad_cantidad = "un.";

                        firstOnly = true;

                        if (fd < new DateTime(2021, 06, 01))
                        {
                            for (int pt = 1; pt <= lista_PS[i].tarifa.numPeriodosTarifarios; pt++)
                            {
                                parcial = CalculaExceso(cal, pt, lista_PS[i]);
                                if (parcial > 0)
                                {
                                    if (firstOnly)
                                    {
                                        facturaDetalle.descripcion = facturaDetalle.descripcion + "AC" + pt + ": "
                                        + string.Format("{0:#,##0.000}", vectorAci[pt]);
                                        firstOnly = false;
                                    }
                                    else
                                    {
                                        facturaDetalle.descripcion = facturaDetalle.descripcion + " AC" + pt + ": "
                                        + string.Format("{0:#,##0.000}", vectorAci[pt]);
                                    }

                                    suma = suma + parcial;
                                }

                            }
                        }
                        else if (fd > new DateTime(2021, 06, 01) && fd < new DateTime(2022, 01, 01))
                        {
                            firstOnly = true;

                            if (curvas.curvaCompleta)
                            {
                                for (int pt = 1; pt <= lista_PS[i].tarifa.numPeriodosTarifarios; pt++)
                                {
                                    parcial = CalculaExceso_anterior_2022_01_01(cal, 
                                        lista_PS[i].tarifa.tarifa, pt, 
                                        lista_PS[i], vectorExcesos);
                                    if (parcial > 0)
                                    {
                                        if (firstOnly)
                                        {
                                            facturaDetalle.descripcion = facturaDetalle.descripcion + "AC" + pt + ": "
                                            + string.Format("{0:#,##0.000}", vectorAci[pt]);
                                            firstOnly = false;
                                        }
                                        else
                                        {
                                            facturaDetalle.descripcion = facturaDetalle.descripcion + " AC" + pt + ": "
                                           + string.Format("{0:#,##0.000}", vectorAci[pt]);
                                        }
                                        

                                        suma = suma + parcial;
                                    }
                                }
                            }
                            else
                            {
                                suma = peaje.importe_excesos_potencia;
                            }



                        }
                        else // Para calculos a partir del 2022_01_01
                        {
                            firstOnly = true;

                            if (curvas.curvaCompleta)
                            {
                                for (int pt = 1; pt <= lista_PS[i].tarifa.numPeriodosTarifarios; pt++)
                                {
                                    parcial = CalculaExceso_posterior_2022_01_01(cal, 
                                        lista_PS[i].tarifa.tarifa, pt, 
                                        lista_PS[i], vectorExcesos, 
                                        vectorPotenciasMaximasRegistradas, dias);

                                    if (parcial > 0)
                                    {
                                        if (firstOnly)
                                        {
                                            facturaDetalle.descripcion = facturaDetalle.descripcion + "AC" + pt + ": "
                                            + string.Format("{0:#,##0.000}", vectorAci[pt]);
                                            firstOnly = false;
                                        }
                                        else
                                        {
                                            facturaDetalle.descripcion = facturaDetalle.descripcion + " AC" + pt + ": "
                                           + string.Format("{0:#,##0.000}", vectorAci[pt]);
                                        }
                                        

                                        suma = suma + parcial;
                                    }
                                }
                            }
                            else
                            {
                                suma = peaje.importe_excesos_potencia;
                            }
                        }                           

                        if (suma > 0)
                        {
                            facturaDetalle.precio = suma;
                            facturaDetalle.unidad_precio = "Eur";
                            facturaDetalle.total = suma;
                            importe_excesos_potencia = facturaDetalle.total;
                            factura.base_imponible = factura.base_imponible + facturaDetalle.total;
                            factura.base_imponible_ie = factura.base_imponible_ie + facturaDetalle.total;
                            factura.lista_factura_detalle.Add(facturaDetalle);
                        }
                        
                        #endregion
                                        
                        #region Dto Cliente
                        if(lista_PS[i].dto_te != 0)
                        {
                            numLineaFactura++;
                            facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                            facturaDetalle.linea_factura = numLineaFactura;
                            facturaDetalle.producto = "DTO_TE";
                            facturaDetalle.concepto = "% Dto. Cliente";
                            facturaDetalle.cantidad = lista_PS[i].dto_te;
                            facturaDetalle.unidad_cantidad = "%";
                            facturaDetalle.precio = importe_termino_energia;
                            facturaDetalle.unidad_precio = "Eur";
                            facturaDetalle.descripcion = string.Format("{0:#,##0.00}", lista_PS[i].dto_te) + "% x "
                                + string.Format("{0:#,##0.00}", importe_termino_energia) + " Eur";
                            facturaDetalle.total = importe_termino_energia * (lista_PS[i].dto_te / 100);
                            factura.base_imponible = factura.base_imponible + facturaDetalle.total;
                            factura.base_imponible_ie = factura.base_imponible_ie + facturaDetalle.total;
                            factura.lista_factura_detalle.Add(facturaDetalle);
                        }

                        #endregion

                        #region Servicio Gestion Preferente
                        if (lista_PS[i].servicio_gestion_preferente != 0)
                        {
                            numLineaFactura++;
                            facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                            facturaDetalle.linea_factura = numLineaFactura;
                            facturaDetalle.producto = "SERV_PREF";
                            facturaDetalle.concepto = "Servicio Gestión Prefer.Precio";

                            if (curvas.curvaCompleta)
                            {
                                facturaDetalle.cantidad = Math.Round((curvas.totalEnergiaActiva / 1000), 2);
                                facturaDetalle.descripcion = string.Format("{0:#,##0.00 MWh}", (curvas.totalEnergiaActiva / 1000)) + " x "
                                + string.Format("{0:#,##0.00 Eur/MWh}", lista_PS[i].servicio_gestion_preferente);
                                facturaDetalle.total = (curvas.totalEnergiaActiva / 1000) * (lista_PS[i].servicio_gestion_preferente);
                            }                                
                            else
                            {
                                facturaDetalle.cantidad = Math.Round((peaje.total_energia_activa / 1000), 2);
                                facturaDetalle.descripcion = string.Format("{0:#,##0.00 MWh}", (peaje.total_energia_activa / 1000)) + " x "
                                + string.Format("{0:#,##0.00 Eur/MWh}", lista_PS[i].servicio_gestion_preferente);
                                facturaDetalle.total = (peaje.total_energia_activa / 1000) * (lista_PS[i].servicio_gestion_preferente);
                            }
                                

                            facturaDetalle.unidad_cantidad = "MWh";
                            facturaDetalle.precio = lista_PS[i].servicio_gestion_preferente;
                            facturaDetalle.unidad_precio = "Eur/MWh";
                            
                            factura.base_imponible = factura.base_imponible + facturaDetalle.total;
                            factura.base_imponible_ie = factura.base_imponible_ie + facturaDetalle.total;
                            factura.lista_factura_detalle.Add(facturaDetalle);
                        }

                        #endregion

                        #region Linea Regularizacion
                        List<EndesaEntity.facturacion.Regularizacion> o;
                        if(regPrecios.dic.TryGetValue(factura.cupsree, out o))
                        {
                            for(int v = 0; v < o.Count; v++)
                            {
                                numLineaFactura++;
                                facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                                facturaDetalle.linea_factura = numLineaFactura;
                                facturaDetalle.cantidad = 1;
                                facturaDetalle.producto = "REG_PRECIOS";
                                facturaDetalle.concepto = o[v].texto_factura;
                                facturaDetalle.descripcion = o[v].texto_factura;
                                facturaDetalle.unidad_precio = "Eur";
                                facturaDetalle.total = o[v].importe;

                                factura.base_imponible = factura.base_imponible + facturaDetalle.total;
                                factura.base_imponible_ie = factura.base_imponible_ie + facturaDetalle.total;
                                factura.lista_factura_detalle.Add(facturaDetalle);
                            }
                        }
                        #endregion

                        #region Alquiler
                        if (lista_PS[i].alquiler != 0)
                        {
                            numLineaFactura++;
                            facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                            facturaDetalle.linea_factura = numLineaFactura;
                            facturaDetalle.producto = "ALQU";
                            facturaDetalle.concepto = "Alquiler Equipos de Medida";
                            facturaDetalle.cantidad = ((lista_PS[i].fecha_fin - lista_PS[i].fecha_inicio).TotalDays + 1);
                            facturaDetalle.unidad_cantidad = "Días";                            
                            facturaDetalle.precio = lista_PS[i].alquiler;
                            facturaDetalle.unidad_cantidad = "Precio/Día";
                            facturaDetalle.total = ((lista_PS[i].fecha_fin - lista_PS[i].fecha_inicio).TotalDays + 1) * lista_PS[i].alquiler;
                            factura.lista_factura_detalle.Add(facturaDetalle);
                            factura.base_imponible = factura.base_imponible + facturaDetalle.total;
                        }
                        #endregion

                        #region Impuesto sobre la electricidad

                        double ise = Convert.ToDouble(param.GetValue("ISE", DateTime.Now, DateTime.Now));
                        double total_energia_activa_MW = 0;
                        double cim = 0;


                        numLineaFactura++;
                        facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                        facturaDetalle.linea_factura = numLineaFactura;
                        facturaDetalle.producto = "IE";
                        facturaDetalle.concepto = "Impuesto sobre la Electricidad";
                        facturaDetalle.cantidad = 1;
                        facturaDetalle.unidad_cantidad = "un.";
                        facturaDetalle.precio = factura.base_imponible_ie;
                        facturaDetalle.unidad_precio = "Eur";
                        //facturaDetalle.descripcion = "5,11269632 % sobre " + string.Format("{0:#,##0.00}", factura.base_imponible_ie) + " Eur";
                        facturaDetalle.descripcion = param.GetValue("ISE", DateTime.Now, DateTime.Now).Replace(".", ",") + " % sobre "
                                + string.Format("{0:#,##0.00}", factura.base_imponible_ie) + " Eur";
                        //facturaDetalle.total = Math.Round(factura.base_imponible_ie * 0.0511269632,2);
                        facturaDetalle.total = Math.Round(factura.base_imponible_ie * (ise / 100), 2);

                        // Calculamos el CIM (cuota integra minima)
                        // Si CIM > IE --> CIM

                        if (curvas.curvaCompleta)
                        {
                            for (int pt = 1; pt < cal.energiaActivaPorPeriodo.Count(); pt++)
                                total_energia_activa_MW = total_energia_activa_MW + cal.energiaActivaPorPeriodo[pt];
                        }
                        else
                        {
                            for (int pt = 1; pt <= lista_PS[i].tarifa.numPeriodosTarifarios; pt++)
                                total_energia_activa_MW = total_energia_activa_MW + peaje.activa[pt];
                        }
                        

                        total_energia_activa_MW = total_energia_activa_MW / 1000;

                        cim = total_energia_activa_MW * 0.5;
                        
                        if(cim > facturaDetalle.total)
                        {
                            facturaDetalle.precio = cim;
                            facturaDetalle.total = cim;
                            facturaDetalle.descripcion = "Cuota Íntegra Mínima sobre I.E";
                        }

                        
                        factura.base_imponible = factura.base_imponible + facturaDetalle.total;
                        factura.lista_factura_detalle.Add(facturaDetalle);


                        #endregion

                        #region Reduccion ISE
                        if (lista_PS[i].exencion_ise)
                        {
                            numLineaFactura++;
                            facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                            facturaDetalle.linea_factura = numLineaFactura;
                            facturaDetalle.producto = "ISE";
                            facturaDetalle.concepto = "Reducción ISE";
                            facturaDetalle.cantidad = 1;
                            facturaDetalle.unidad_cantidad = "un.";
                            facturaDetalle.precio = ((factura.base_imponible_ie * 0.85) * 
                                (lista_PS[i].porcentaje_exencion / 100)) * -1;
                            facturaDetalle.unidad_precio = "Eur";
                            //facturaDetalle.descripcion = "-5,11269632 % sobre " + string.Format("{0:#,##0.00}", Math.Abs(facturaDetalle.precio)) + " Eur";
                            facturaDetalle.descripcion = "-" +
                                param.GetValue("ISE", DateTime.Now, DateTime.Now).Replace(".",",") + " % sobre " 
                                + string.Format("{0:#,##0.00}", Math.Abs(facturaDetalle.precio)) + " Eur";
                            //facturaDetalle.total = Math.Round(facturaDetalle.precio * 0.0511269632, 2);
                            facturaDetalle.total = Math.Round(facturaDetalle.precio * 
                                (ise / 100), 2);

                            // Si se aplica el CIM no debe tenerse en cuenta la reducción ISE
                            if (cim < facturaDetalle.total)
                            {
                                factura.base_imponible = factura.base_imponible + facturaDetalle.total;
                            }
                            
                            factura.lista_factura_detalle.Add(facturaDetalle);
                        }
                        #endregion

                        #region IVA
                        numLineaFactura++;
                        facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                        facturaDetalle.linea_factura = numLineaFactura;
                        facturaDetalle.producto = "IVA";
                        facturaDetalle.concepto = "IVA Normal";
                        facturaDetalle.descripcion = param.GetValue("IVA_Normal") + " % sobre " 
                            + string.Format("{0:#,##0.00}", factura.base_imponible)
                                + " Eur";
                        facturaDetalle.total = Math.Round(factura.base_imponible * 
                            (Convert.ToDouble(param.GetValue("IVA_Normal", DateTime.Now, DateTime.Now)) / 100), 2);
                        factura.iva = facturaDetalle.total;
                        factura.lista_factura_detalle.Add(facturaDetalle);
                        #endregion

                        factura.total_factura = factura.base_imponible * 
                            (1 + (Convert.ToDouble(param.GetValue("IVA_Normal", DateTime.Now, DateTime.Now)) / 100));

                        #region Texto
                        numLineaFactura++;
                        facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                        facturaDetalle.linea_factura = numLineaFactura;
                        facturaDetalle.producto = "texto";

                        if (curvas.curvaCompleta)
                        {

                            if (fd >= new DateTime(2021, 06, 01))
                            {
                                if (fd >= new DateTime(2022, 01, 01) && fh <= new DateTime(2022,03,30).Date)
                                    facturaDetalle.descripcion = preciosRegulados.Texto_Componentes_Regulados_2022(total_TP + importe_excesos_potencia,
                                total_TE, 
                                total_ATR_Cargo_Potencia, total_ATR_Cargo_Energia,
                                lista_PS[i].tarifa,
                                cal.energiaActivaPorPeriodo, importe_energia_reactiva);
                                else if (fd >= new DateTime(2022, 03, 31).Date)
                                    facturaDetalle.descripcion = preciosRegulados.Texto_Componentes_Regulados_20220331(total_TP + importe_excesos_potencia,
                                total_TE,
                                total_ATR_Cargo_Potencia, total_ATR_Cargo_Energia,
                                lista_PS[i].tarifa,
                                cal.energiaActivaPorPeriodo, importe_energia_reactiva);
                                else
                                    facturaDetalle.descripcion = preciosRegulados.Texto_Componentes_Regulados_2021(total_TP + importe_excesos_potencia,
                                total_TE,
                                total_ATR_Cargo_Potencia, total_ATR_Cargo_Energia,
                                lista_PS[i].tarifa,
                                cal.energiaActivaPorPeriodo, importe_energia_reactiva);
                            }
                            else
                            {
                                facturaDetalle.descripcion = preciosRegulados.Texto_Componentes_Regulados(importe_potencias,
                                importe_excesos_potencia, lista_PS[i].tarifa,
                                cal.energiaActivaPorPeriodo, importe_energia_reactiva);
                            }

                            
                        }
                        else
                        {
                            if (fd >= new DateTime(2021, 06, 01))
                            {
                                if (fd >= new DateTime(2022, 01, 01) && fh <= new DateTime(2022, 03, 30).Date)
                                    facturaDetalle.descripcion = 
                                    preciosRegulados.Texto_Componentes_Regulados_2022(total_TP + importe_excesos_potencia,
                                    total_TE,
                                    total_ATR_Cargo_Potencia, total_ATR_Cargo_Energia,
                                    lista_PS[i].tarifa,
                                    peaje.activa, importe_energia_reactiva);
                                else if (fd >= new DateTime(2022, 03, 31).Date)
                                    facturaDetalle.descripcion = preciosRegulados.Texto_Componentes_Regulados_20220331(total_TP + importe_excesos_potencia,
                                    total_TE,
                                    total_ATR_Cargo_Potencia, total_ATR_Cargo_Energia,
                                    lista_PS[i].tarifa,
                                    cal.energiaActivaPorPeriodo, importe_energia_reactiva);
                                else
                                    facturaDetalle.descripcion =
                                    preciosRegulados.Texto_Componentes_Regulados_2021(total_TP + importe_excesos_potencia,
                                    total_TE, 
                                    total_ATR_Cargo_Potencia, total_ATR_Cargo_Energia,
                                    lista_PS[i].tarifa,
                                    peaje.activa, importe_energia_reactiva);
                            }
                            else
                            {
                                facturaDetalle.descripcion = preciosRegulados.Texto_Componentes_Regulados(importe_potencias,
                                importe_excesos_potencia, lista_PS[i].tarifa,
                                peaje.activa, importe_energia_reactiva);
                            }

                                
                        }


                        //if (param.GetValue("CD_RDL17/2021") == "S")
                        if (fd < new DateTime(2022, 01, 01))
                        {
                            descuento = (total_ATR_Cargo_Potencia + total_ATR_Cargo_Energia)
                                - (total_ATR_Cargo_Potencia_descuento + total_ATR_Cargo_Energia_descuento);

                            numLineaFactura++;
                            factura.lista_factura_detalle.Add(facturaDetalle);                            
                            facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                            facturaDetalle.linea_factura = numLineaFactura;
                            facturaDetalle.producto = "texto_1";
                            facturaDetalle.descripcion = preciosRegulados.Texto_RDL17_2021(Math.Abs(descuento));
                            factura.lista_factura_detalle.Add(facturaDetalle);
                        }
                        else 
                        {
                            numLineaFactura++;
                            factura.lista_factura_detalle.Add(facturaDetalle);
                            facturaDetalle = new EndesaEntity.eer.FacturaDetalle();
                            facturaDetalle.linea_factura = numLineaFactura;
                            facturaDetalle.producto = "texto_1";
                            facturaDetalle.descripcion = preciosRegulados.Texto_Actual();
                            factura.lista_factura_detalle.Add(facturaDetalle);
                        }
                       
                        #endregion

                        fechaHora_gen_factura = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                        // Generamos SSTT
                        //ImprimirHojaSuplementoTerritorial_OfficeOpenXML(factura.cupsree, lista_PS[i].cuotas, fd.ToString("yyyyMM"), fechaHora_gen_factura);

                        // Generamos Factura
                        //GeneralArchivoExcel_OfficeOpenXML(factura, fd, fh, fechaHora_gen_factura);                      


                        GeneraArchivoExcel_MSOffice(factura, lista_PS[i].fecha_inicio, lista_PS[i].fecha_fin, fechaHora_gen_factura,!curvas.curvaCompleta);

                        if (exportar_calculos)
                        {
                            Exportar_Calculos exportarcalculos = new Exportar_Calculos();
                            exportarcalculos.GeneraExcelCalculos(lista_PS[i], curvas, cal, fd, fechaHora_gen_factura,
                                coef_excesos, precio_excesos, vectorExcesos);
                        }
                            

                        factEER.GuardaFactura(factura);
                    }


                }
                else
                {
                    MessageBox.Show("Medida incompleta para el punto: " + lista_PS[i].cups20,
                       "Medida Incompleta",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Error);
                }
                
            }
                                    
        }
       


        private double[] CargaPotenciasMaximasRegistradas(EndesaBusiness.calendarios.Calendario cal, EndesaEntity.punto_suministro.PuntoSuministro ps)
        {

            // Devuelve la potencia máxima registrada por periodo tarifario
            // en todo el periodo de consumo si este es superior a la potencia
            // contratada, en caso contrario devuelve 0.

            int pt = 0;
            double[] potenciasMaximasRegistradas;
            double potencia_contratada = 0;

            potenciasMaximasRegistradas = new double[ps.tarifa.numPeriodosTarifarios + 1];

            for (int i = 1; i <= curvas.numPeriodosMedidaCuartoHorario; i++)
            {
                pt = cal.vectorPeriodosTarifariosCuartoHorarios[i];
                potencia_contratada = ps.potecias_contratadas[pt];

                if ((curvas.curvaCuartoHorariaPotencias[i] > potencia_contratada)
                    && (potenciasMaximasRegistradas[pt] < curvas.curvaCuartoHorariaPotencias[i]))
                    potenciasMaximasRegistradas[pt] = curvas.curvaCuartoHorariaPotencias[i];
            }


            return potenciasMaximasRegistradas;

        }

        private double[] CargaExcesos(EndesaBusiness.calendarios.Calendario cal, EndesaEntity.punto_suministro.PuntoSuministro ps)
        {
            int pt = 0;
            
            double potencia_contratada = 0;
            double diferencia = 0;

            double[] vectorExcesos = new double[(curvas.numPeriodosMedidaHorario + 1) * 4];

            for (int i = 1; i <= curvas.numPeriodosMedidaCuartoHorario; i++)
            {
                pt = cal.vectorPeriodosTarifariosCuartoHorarios[i];
                potencia_contratada = ps.potecias_contratadas[pt];

                if (curvas.curvaCuartoHorariaPotencias[i] > potencia_contratada)
                {
                    diferencia = curvas.curvaCuartoHorariaPotencias[i] - potencia_contratada;
                    vectorExcesos[i] = vectorExcesos[i] + diferencia;
                    vectorExcesos[i] = (vectorExcesos[i] * vectorExcesos[i]);
                }


            }
            return vectorExcesos;
        }

        private double CalculaExceso(EndesaBusiness.calendarios.Calendario cal, int periodoTarifario, 
            EndesaEntity.punto_suministro.PuntoSuministro ps)
        {
            double total = 0;
            double suma = 0;


            double[] vectorExcesos = CargaExcesos(cal, ps);
            double constanteExcesos = Convert.ToDouble(param.GetValue("constanteExcesos", DateTime.Now, DateTime.Now));

            for(int i = 1; i < cal.numPeriodosMedidaCuartoHorario; i++)            
                if (cal.vectorPeriodosTarifariosCuartoHorarios[i] == periodoTarifario)
                    suma = suma + vectorExcesos[i];

            if(suma > 0)
            {
                total = Math.Sqrt(suma) * coef_excesos.GetValorExcesosPotencia("global", periodoTarifario) * Math.Round(constanteExcesos / 166.386, 4);
                if (ps.tarifa.tarifa.Substring(0, 1) != "2" &&
                    ps.tarifa.tarifa.Substring(0, 1) != "3")
                    vectorAci[periodoTarifario] = Math.Sqrt(suma);
            }
            else
            {
                total = 0;
                // vectorAci[periodoTarifario] = 0;
            }


            return total;
        }



        private double CalculaExceso_anterior_2022_01_01(EndesaBusiness.calendarios.Calendario cal, string tarifa, int periodoTarifario,
            EndesaEntity.punto_suministro.PuntoSuministro ps, double[] vectorExcesos)
        {
            double total = 0;
            double suma = 0;
            
            double constanteExcesos = Convert.ToDouble(param.GetValue("constanteExcesos", DateTime.Now, DateTime.Now));


            for (int i = 1; i < cal.numPeriodosMedidaCuartoHorario; i++)
                if (cal.vectorPeriodosTarifariosCuartoHorarios[i] == periodoTarifario)
                    suma = suma + vectorExcesos[i];

            if (suma > 0)
            {
                total = Math.Sqrt(suma) * coef_excesos.GetValorExcesosPotencia(tarifa, periodoTarifario) * Math.Round(constanteExcesos / 166.386, 4);
                if (ps.tarifa.tarifa.Substring(0, 1) != "2")                    
                    vectorAci[periodoTarifario] = Math.Sqrt(suma);
            }
            else
            {
                total = 0;                
            }


            return total;
        }

        private double CalculaExceso_posterior_2022_01_01(EndesaBusiness.calendarios.Calendario cal, 
            string tarifa, int periodoTarifario, EndesaEntity.punto_suministro.PuntoSuministro ps,
            double[] vectorExcesos, double[] vectorPotenciasMaximasRegistradas, int diasFactura)
        {
            double total = 0;
            double suma = 0;
            double kp = 0;
            double tep = 0;

            kp = coef_excesos.GetValorExcesosPotencia(tarifa, periodoTarifario);
            tep = precio_excesos.GetValorPrecioExcesosPotencia(tarifa, ps.tipo_punto_medida);

            if (ps.tipo_punto_medida < 4)
            {
                vectorExcesos = CargaExcesos(cal, ps);

                for (int i = 1; i < cal.numPeriodosMedidaCuartoHorario; i++)
                    if (cal.vectorPeriodosTarifariosCuartoHorarios[i] == periodoTarifario)
                        suma = suma + vectorExcesos[i];

                if (suma > 0)
                {
                    total = kp * tep * Math.Sqrt(suma);
                    vectorAci[periodoTarifario] = Math.Sqrt(suma);
                }
                else
                {
                    total = 0;
                }
            }
            else
            {
                vectorPotenciasMaximasRegistradas = CargaPotenciasMaximasRegistradas(cal, ps);
                if (vectorPotenciasMaximasRegistradas[periodoTarifario] > 0)
                    total =
                        2 * (vectorPotenciasMaximasRegistradas[periodoTarifario] - ps.potecias_contratadas[periodoTarifario])
                        * tep * diasFactura;
                vectorAci[periodoTarifario] = total;

            }
                            

            return total;
        }


        private void GeneraArchivoExcel_MSOffice(EndesaEntity.eer.Factura factura, DateTime fd, DateTime fh, string fechaHoraGen, bool desdePeajes)
        {
            DirectoryInfo dirSalida;
            string nombreFicheroFactura = "";
            bool hay_sstt = false;
            utilidades.Fechas ff = new utilidades.Fechas();
            double meses = CoeficienteMes(fd, fh);

            int c = 0;
            int f = 0;


            int dias = 0;
            int dias_anio = 0;

            #region Variables Factura
            double total_importe_energia_variable = 0;
            double total_importe_potencia = 0;
            #endregion

            utilidades.Fechas utilFechas = new utilidades.Fechas();

            dias = Convert.ToInt32((fh - fd).TotalDays + 1);
            dias_anio = utilFechas.EsAnioBisiesto(fd) ? 366 : 365;


            // Ruta de la plantilla de facturacion
            FileInfo fichero = new FileInfo(param.GetValue("ruta_plantilla_factura", DateTime.Now, DateTime.Now)
                + factura.cupsree + ".xlsx");

            dirSalida = new DirectoryInfo(param.GetValue("ruta_salida_facturas", DateTime.Now, DateTime.Now));

            if (!dirSalida.Exists)
                dirSalida.Create();

            nombreFicheroFactura = factura.cupsree + "_" + fd.ToString("yyyyMM") + "_" + fechaHoraGen + ".xlsx";

            FileInfo filesave = new FileInfo(dirSalida.FullName + "\\" + nombreFicheroFactura);
            office.Excel excel = new office.Excel(fichero.FullName);


            excel.PonValor(28, 5, factura.fecha_consumo_desde.ToString("dd/MM/yyyy") + " al " + factura.fecha_consumo_hasta.ToString("dd/MM/yyyy"));


            #region Termino Energia Variable
            f = 40;
            c = 4;
            List<EndesaEntity.eer.FacturaDetalle> lineas_energia_variable = factura.lista_factura_detalle.FindAll(z => z.producto == "L01");
            for (int i = 0; i < lineas_energia_variable.Count; i++)
            {
                total_importe_energia_variable = total_importe_energia_variable + lineas_energia_variable[i].total;
                excel.PonValor(f, 6, lineas_energia_variable[i].descripcion);
                f++;
            }

            // excel.PonValor(39, 12, string.Format("{0:#,##0.00}", total_importe_energia_variable));
            excel.PonValor(39, 12, total_importe_energia_variable);
            excel.EstiloMillares(39, 12);

            #endregion

            #region Facturacion Potencia Periodos
            f = 49;
            List<EndesaEntity.eer.FacturaDetalle> lineas_potencia;

            if (fd < new DateTime(2021, 06, 01))
            {
                lineas_potencia = factura.lista_factura_detalle.FindAll(z => z.producto == "L34");
                for (int i = 0; i < lineas_potencia.Count; i++)
                {
                    total_importe_potencia = total_importe_potencia + lineas_potencia[i].total;
                    if (!desdePeajes)
                    {
                        excel.PonValor(f, 6, lineas_potencia[i].descripcion);
                        f++;
                    }
                }
            }
            else
            {
                lineas_potencia = factura.lista_factura_detalle.FindAll(z => z.producto == "L85");
                for (int i = 0; i < lineas_potencia.Count; i++)
                {
                    total_importe_potencia = total_importe_potencia + lineas_potencia[i].total;
                    //if (!desdePeajes)
                    //{
                    //    excel.PonValor(f, 6, lineas_potencia[i].descripcion);
                    //    f++;
                    //}

                    excel.PonValor(f, 6, lineas_potencia[i].descripcion);
                    f++;
                }
            }

            if (fd < new DateTime(2021, 06, 01))
            {
                if (!desdePeajes)
                    excel.PonValor(f, 6, string.Format("{0:#,##0.00}", (total_importe_potencia))
                    + " Eur x " + meses + " MESES / 12 MESES");

                if (!desdePeajes)
                    excel.PonValor(48, 12, (meses * total_importe_potencia) / 12);
                else
                    excel.PonValor(48, 12, total_importe_potencia);

                excel.EstiloMillares(48, 12);
            }
            else
            {
                //if (!desdePeajes)
                //    excel.PonValor(f, 6, string.Format("{0:#,##0.00}", (total_importe_potencia))
                //    + " €/Año x " + dias + " DÍAS / " + dias_anio + " DIAS");

                //if (!desdePeajes)
                //    excel.PonValor(48, 12, (dias * total_importe_potencia) / dias_anio);
                //else
                //    excel.PonValor(48, 12, total_importe_potencia);

                excel.PonValor(f, 6, string.Format("{0:#,##0.00}", (total_importe_potencia))
                    + " €/Año x " + dias + " DÍAS / " + dias_anio + " DIAS");
                excel.PonValor(48, 12, (dias * total_importe_potencia) / dias_anio);



                excel.EstiloMillares(48, 12);
            }
            


            f++;
            #endregion

            #region Excesos Potencia
            List<EndesaEntity.eer.FacturaDetalle> linea_excesos_potencia = factura.lista_factura_detalle.FindAll(z => z.producto == "REPA");
            for (int i = 0; i < linea_excesos_potencia.Count; i++)
            {
                excel.PonValor(f, 3, "Recargo por Excesos de Potencia");
                excel.PonValor(f, 8, linea_excesos_potencia[i].descripcion);
                excel.PonValor(f, 12, linea_excesos_potencia[i].total);
                excel.EstiloMillares(f, 12);

                f++;
            }
            #endregion

            #region Excesos Reactiva
            if (fd < new DateTime(2021, 06, 01))
            {
                List<EndesaEntity.eer.FacturaDetalle> linea_reactiva = factura.lista_factura_detalle.FindAll(z => z.producto == "REAC");
                for (int i = 0; i < linea_reactiva.Count; i++)
                {
                    excel.PonValor(f, 3, "Energía Reactiva");
                    excel.PonValor(f, 12, linea_reactiva[i].total);
                    excel.EstiloMillares(f, 12);

                    f++;
                }
            }
            else
            {
                List<EndesaEntity.eer.FacturaDetalle> linea_reactiva = factura.lista_factura_detalle.FindAll(z => z.producto == "L44");
                for (int i = 0; i < linea_reactiva.Count; i++)
                {
                    excel.PonValor(f, 3, "Energía Reactiva");
                    excel.PonValor(f, 12, linea_reactiva[i].total);
                    excel.EstiloMillares(f, 12);

                    f++;
                }
            }
               

            #endregion

            #region Dto Cliente
            List<EndesaEntity.eer.FacturaDetalle> dto_cliente = factura.lista_factura_detalle.FindAll(z => z.producto == "DTO_TE");
            for (int i = 0; i < dto_cliente.Count; i++)
            {
                excel.PonValor(f, 3, "% Dto. Cliente");
                excel.PonValor(f, 8, dto_cliente[i].descripcion);
                excel.PonValor(f, 12,  dto_cliente[i].total);
                excel.EstiloMillares(f, 12);

                f++;
            }
            #endregion

            #region Servicio Gestion Preferente
            List<EndesaEntity.eer.FacturaDetalle> gest_pref = factura.lista_factura_detalle.FindAll(z => z.producto == "SERV_PREF");
            for (int i = 0; i < gest_pref.Count; i++)
            {
                excel.PonValor(f, 3, "Servicio Gestión Prefer. Precio");
                excel.PonValor(f, 8, gest_pref[i].descripcion);
                excel.PonValor(f, 12,  gest_pref[i].total);
                excel.EstiloMillares(f, 12);

                f++;
            }

            #endregion

            #region SSTT
            List<EndesaEntity.eer.FacturaDetalle> lineas_sstt = factura.lista_factura_detalle.FindAll(z => z.producto == "SSTT");
            for (int i = 0; i < lineas_sstt.Count; i++)
            {
                hay_sstt = true;
                excel.PonValor(f, 3, "Suplemento Territorial de Peaje (I)");
                excel.PonValor(f, 12,  lineas_sstt[i].total);
                excel.EstiloMillares(f, 12);

                f++;
            }
            #endregion

            #region Regularizacion
            List<EndesaEntity.eer.FacturaDetalle> lineas_reg = factura.lista_factura_detalle.FindAll(z => z.producto == "REG_PRECIOS");
            f = 67;
            for (int i = 0; i < lineas_reg.Count; i++)
            {                
                excel.PonValor(f, 3, lineas_reg[i].descripcion);
                excel.PonValor(f, 12, lineas_reg[i].total);
                excel.EstiloMillares(f, 12);

                f++;
            }
            #endregion



            #region Impuesto sobre la electricidad
            f = 68;
            List<EndesaEntity.eer.FacturaDetalle> linea_ise = factura.lista_factura_detalle.FindAll(z => z.producto == "IE");
            for (int i = 0; i < linea_ise.Count; i++)
            {
                excel.PonValor(f, 3, "Impuesto sobre la Electricidad");
                excel.PonValor(f, 8, linea_ise[i].descripcion);
                excel.PonValor(f, 12,  linea_ise[i].total);
                excel.EstiloMillares(f, 12);

                f++;
            }

            #endregion

            #region Reduccion ISE
            f = 69;
            List<EndesaEntity.eer.FacturaDetalle> linea_exencion_ise = factura.lista_factura_detalle.FindAll(z => z.producto == "ISE");
            for (int i = 0; i < linea_exencion_ise.Count; i++)
            {
                excel.PonValor(f, 3, "Reducción ISE");
                excel.PonValor(f, 8, linea_exencion_ise[i].descripcion);
                excel.PonValor(f, 12, linea_exencion_ise[i].total);
                excel.EstiloMillares(f, 12);

                f++;
            }

            #endregion


            #region Alquiler
            List<EndesaEntity.eer.FacturaDetalle> linea_alquiler = factura.lista_factura_detalle.FindAll(z => z.producto == "ALQU");
            for (int i = 0; i < linea_alquiler.Count; i++)
            {
                excel.PonValor(f, 3, "Alquiler de Equipos de Medida");
                excel.PonValor(f, 12,  linea_alquiler[i].total);
                excel.EstiloMillares(f, 12);
                f++;
            }
            #endregion


            #region Iva normal
            List<EndesaEntity.eer.FacturaDetalle> linea_iva = factura.lista_factura_detalle.FindAll(z => z.producto == "IVA");
            for (int i = 0; i < linea_iva.Count; i++)
            {
                excel.PonValor(f, 3, "IVA normal");
                excel.PonValor(f, 8, linea_iva[i].descripcion);
                excel.PonValor(f, 12,  linea_iva[i].total);
                excel.EstiloMillares(f, 12);
                f++;
            }

            #endregion

            #region Totales
            excel.PonValor(31, 7, string.Format("{0:#,##0.00}", factura.total_factura) + " €");
            excel.PonValor(87, 12, factura.total_factura);
            excel.EstiloMillares(87, 12);



            //if (hay_sstt)
            //{
            //    excel.PonValor(87, 12, factura.total_factura);
            //    excel.EstiloMillares(87, 12);
            //}
            //else
            //{
            //    excel.PonValor(90, 12,  factura.total_factura);
            //    excel.EstiloMillares(90, 12);

            //}


            #endregion

            #region Texto RDL17/2021
            if (fd < new DateTime(2022, 01, 01))
            {
                List<EndesaEntity.eer.FacturaDetalle> linea_texto_1 = factura.lista_factura_detalle.FindAll(z => z.producto == "texto_1");
                for (int i = 0; i < linea_texto_1.Count; i++)
                {
                    excel.PonValor(90, 3, linea_texto_1[i].descripcion);
                }
            }
               
            #endregion

            #region Texto coste peajes
            List<EndesaEntity.eer.FacturaDetalle> linea_texto = factura.lista_factura_detalle.FindAll(z => z.producto == "texto");
            for (int i = 0; i < linea_texto.Count; i++)
            {
                excel.PonValor(93, 3, linea_texto[i].descripcion);
            }

            #endregion

            excel.GuardarComo(filesave.FullName, "xlsx");
            excel.Cerrar();
            excel = null;

        }

        private void GeneralArchivoExcel_OfficeOpenXML(EndesaEntity.eer.Factura factura, DateTime fd, DateTime fh, string fechaHoraGen)
        {
            DirectoryInfo dirSalida;
            string nombreFicheroFactura = "";
            bool hay_sstt = false;

            int c = 0;
            int f = 0;

            #region Variables Factura
            double total_importe_energia_variable = 0;
            double total_importe_potencia = 0;
            #endregion

            

            // Ruta de la plantilla de facturacion
            FileInfo fichero = new FileInfo(param.GetValue("ruta_plantilla_factura", DateTime.Now, DateTime.Now)
                + factura.cupsree + ".xlsx");

            dirSalida = new DirectoryInfo(param.GetValue("ruta_salida_facturas", DateTime.Now, DateTime.Now));
                
            if (!dirSalida.Exists)
                dirSalida.Create();

            nombreFicheroFactura = factura.cupsree + "_" + fd.ToString("yyyyMM") + "_" + fechaHoraGen + ".xlsx";

            FileInfo filesave = new FileInfo(dirSalida.FullName + "\\" + nombreFicheroFactura);


            //FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);    
            ExcelPackage excelPackage = new ExcelPackage(fichero);
            var workSheet = excelPackage.Workbook.Worksheets["Factura"];


            #region Datos del Cliente
            //workSheet.Cells["CuadroTexto 13"].Value = "Prueba";
            //f = 4;
            //workSheet.Cells[4, 8].Value = "       Razón Social: " + factura.nombre_cliente; 
            //workSheet.Cells[5, 8].Value = "       NIF/CIF: " + factura.nif; 
            //workSheet.Cells[6, 8].Value = "       Dir. Fiscal: " + factura.direccion_facturacion; 
            //workSheet.Cells[8, 8].Value = "       Dir. Suministro: " + factura.direccion_suministro; 
            //workSheet.Cells[10, 8].Value = "       CUPS: " + factura.cupsree + "0F"; 

            //f = 17;
            //workSheet.Cells[f, 8].Value = "          " + factura.nombre_cliente; f++;
            #endregion

            #region Resumen de la factura

            #endregion

            


            workSheet.Cells[28, 5].Value = fd.ToString("dd/MM/yyyy") + " al " + fh.ToString("dd/MM/yyyy");


            #region Termino Energia Variable
            f = 40;
            c = 4;
            List<EndesaEntity.eer.FacturaDetalle> lineas_energia_variable = factura.lista_factura_detalle.FindAll(z => z.producto == "L01");
            for(int i = 0; i < lineas_energia_variable.Count; i++)
            {
                total_importe_energia_variable = total_importe_energia_variable + lineas_energia_variable[i].total;
                workSheet.Cells[f, 6].Value = lineas_energia_variable[i].descripcion;                
                f++;
            }

            workSheet.Cells[39, 12].Value = total_importe_energia_variable;
            workSheet.Cells[39, 12].Style.Numberformat.Format = "#,##0.00";


            #endregion

            #region Facturacion Potencia Periodos
            f = 49;
            List<EndesaEntity.eer.FacturaDetalle> lineas_potencia = factura.lista_factura_detalle.FindAll(z => z.producto == "L34");
            for (int i = 0; i < lineas_potencia.Count; i++)
            {
                total_importe_potencia = total_importe_potencia + lineas_potencia[i].total;
                workSheet.Cells[f, 6].Value = lineas_potencia[i].descripcion;
                f++;
            }

            workSheet.Cells[f, 6].Value = string.Format("{0:#,##0}", total_importe_potencia)
                + " Eur x 1 MESES / 12 MESES";

            workSheet.Cells[48, 12].Value = total_importe_potencia / 12;
            workSheet.Cells[48, 12].Style.Numberformat.Format = "#,##0.00";
            f++;
            #endregion

            #region Excesos Potencia
            List<EndesaEntity.eer.FacturaDetalle> linea_excesos_potencia = factura.lista_factura_detalle.FindAll(z => z.producto == "REPA");
            for (int i = 0; i < linea_excesos_potencia.Count; i++)
            {
                workSheet.Cells[f, 3].Value = "Recargo por Excesos de Potencia";
                workSheet.Cells[f, 8].Value = linea_excesos_potencia[i].descripcion;
                workSheet.Cells[f, 12].Value = linea_excesos_potencia[i].total;
                workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                f++;
            }
            #endregion

            #region Excesos Reactiva
            List<EndesaEntity.eer.FacturaDetalle> linea_reactiva = factura.lista_factura_detalle.FindAll(z => z.producto == "REAC");
            for (int i = 0; i < linea_reactiva.Count; i++)
            {
                workSheet.Cells[f, 3].Value = "Energía Reactiva";
                workSheet.Cells[f, 12].Value = linea_reactiva[i].total;
                workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                f++;
            }

            #endregion

            #region Dto Cliente
            List<EndesaEntity.eer.FacturaDetalle> dto_cliente = factura.lista_factura_detalle.FindAll(z => z.producto == "DTO_TE");
            for (int i = 0; i < dto_cliente.Count; i++)
            {
                workSheet.Cells[f, 3].Value = "% Dto. Cliente";
                workSheet.Cells[f, 8].Value = dto_cliente[i].descripcion;
                workSheet.Cells[f, 12].Value = dto_cliente[i].total;
                workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                f++;
            }
            #endregion

            #region Servicio Gestion Preferente
            List<EndesaEntity.eer.FacturaDetalle> gest_pref = factura.lista_factura_detalle.FindAll(z => z.producto == "SERV_PREF");
            for (int i = 0; i < gest_pref.Count; i++)
            {
                workSheet.Cells[f, 3].Value = "Servicio Gestión Prefer. Precio";
                workSheet.Cells[f, 8].Value = gest_pref[i].descripcion;
                workSheet.Cells[f, 12].Value = gest_pref[i].total;
                workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                f++;
            }

            #endregion

            #region SSTT
            List<EndesaEntity.eer.FacturaDetalle> lineas_sstt = factura.lista_factura_detalle.FindAll(z => z.producto == "SSTT");
            for (int i = 0; i < lineas_sstt.Count; i++)
            {
                hay_sstt = true;
                workSheet.Cells[f, 3].Value = "Suplemento Territorial de Peaje (I)";
                workSheet.Cells[f, 12].Value = lineas_sstt[i].total;
                workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                f++;
            }
            #endregion
                     

            #region Impuesto sobre la electricidad
            f = 65;
            List<EndesaEntity.eer.FacturaDetalle> linea_ise = factura.lista_factura_detalle.FindAll(z => z.producto == "IE");
            for (int i = 0; i < linea_ise.Count; i++)
            {
                workSheet.Cells[f, 3].Value = "Impuesto sobre la Electricidad";
                workSheet.Cells[f, 8].Value = linea_ise[i].descripcion;
                workSheet.Cells[f, 12].Value = linea_ise[i].total;
                workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                f++;
            }


            #endregion

            #region Alquiler
            List<EndesaEntity.eer.FacturaDetalle> linea_alquiler = factura.lista_factura_detalle.FindAll(z => z.producto == "ALQU");
            for (int i = 0; i < linea_alquiler.Count; i++)
            {
                workSheet.Cells[f, 3].Value = "Alquiler de Equipos de Medida";                
                workSheet.Cells[f, 12].Value = linea_alquiler[i].total;
                workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                f++;
            }
            #endregion


            #region Iva normal
            List<EndesaEntity.eer.FacturaDetalle> linea_iva = factura.lista_factura_detalle.FindAll(z => z.producto == "IVA");
            for (int i = 0; i < linea_iva.Count; i++)
            {
                workSheet.Cells[f, 3].Value = "IVA normal";
                workSheet.Cells[f, 8].Value = linea_iva[i].descripcion;
                workSheet.Cells[f, 12].Value = linea_iva[i].total;
                workSheet.Cells[f, 12].Style.Numberformat.Format = "#,##0.00";
                f++;
            }
                       
            #endregion

            #region Totales
            workSheet.Cells[31, 7].Value = factura.total_factura;
            workSheet.Cells[31, 7].Style.Numberformat.Format = "#,##0.00 €";

            if (hay_sstt)
            {
                workSheet.Cells[87, 12].Value = factura.total_factura;
                workSheet.Cells[87, 12].Style.Numberformat.Format = "#,##0.00";
            }
            else
            {
                workSheet.Cells[90, 12].Value = factura.total_factura;
                workSheet.Cells[90, 12].Style.Numberformat.Format = "#,##0.00";
            }

            
            #endregion

            #region Texto coste peajes
            List<EndesaEntity.eer.FacturaDetalle> linea_texto = factura.lista_factura_detalle.FindAll(z => z.producto == "texto");
            for (int i = 0; i < linea_texto.Count; i++)
            {
                workSheet.Cells[96, 3].Value = linea_texto[i].descripcion;
            }
                
            #endregion


            excelPackage.SaveAs(filesave);
            excelPackage = null;


        }

        
        public void ImprimirHojaSuplementoTerritorial_OfficeOpenXML(string cups20, int cuotas_totales, string yyyymm, string fechaHoraGen)
        {
            int c = 0;
            int f = 0;
            string cabecera = "(I) Suplemento Territorial de peaje: Datos de cálculo";
            string pie1 = "Número de facturas de regularización: ";
            string nombreFichero = "";
            try
            {

                nombreFichero = cups20 + "_SSTT_" + yyyymm + "_" + fechaHoraGen + ".xlsx";

                FileInfo filesave = new FileInfo(param.GetValue("ruta_salida_facturas", DateTime.Now, DateTime.Now) + "\\" + nombreFichero);
                ExcelPackage excelPackage = new ExcelPackage(filesave);
                var workSheet = excelPackage.Workbook.Worksheets.Add("SSTT");

                f = 3;
                c = 2;

                workSheet.Cells[f, c].Value = cabecera; 
                workSheet.Cells[f, c].Style.Font.Bold = true;
                f++;
                List<string> lista = suplementoTerritorial.GetComplementoTerritorial(cups20);
                for(int i = 0; i < lista.Count(); i++)
                {
                    workSheet.Cells[f, c].Value = lista[i]; f++;
                }
                f++;
                workSheet.Cells[f, c].Value = pie1 + cuotas_totales; f++;
                f++;
                workSheet.Cells[f, c].Value = "Número de fracciones de regularización: "
                    + (cuotas_totales - suplementoTerritorial.CuotasPdtes(cups20, Convert.ToInt32(yyyymm))).ToString() + " de " + cuotas_totales;


                excelPackage.SaveAs(filesave);
                excelPackage = null;

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }        
        }

        public void ImprimirHojaSuplementoTerritorial_MSOffice(string cups20, int cuotas_totales, string yyyymm)
        {
            int c = 0;
            int f = 0;
            string cabecera = "(I) Suplemento Territorial de peaje: Datos de cálculo";
            string pie1 = "Número de facturas de regularización: ";
            string nombreFichero = "";

           
            try
            {

                FileInfo filesave = new FileInfo(param.GetValue("ruta_salida_facturas", DateTime.Now, DateTime.Now) + "\\" + nombreFichero);
                office.Excel excel = new office.Excel();
                
                f = 3;
                c = 2;

                excel.PonValor(f,c, cabecera); f++;

                List<string> lista = suplementoTerritorial.GetComplementoTerritorial(cups20);
                for (int i = 0; i < lista.Count(); i++)
                {
                    excel.PonValor(f, c, lista[i]); f++;
                }
                f++;
                excel.PonValor(f, c, pie1 + cuotas_totales); f++;
                f++;
                excel.PonValor(f, c, "Número de fracciones de regularización: "
                    + (cuotas_totales - suplementoTerritorial.CuotasPdtes(cups20, Convert.ToInt32(yyyymm))).ToString() + " de " + cuotas_totales);
                               
                nombreFichero = cups20 + "_SSTT_" + yyyymm + ".xlsx";
                // excel.GuardarComo(filesave.FullName,)

                

                

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }


        private int UltimaFactura()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            int ultimaFactura = 0;
            try
            {
                strSql = "select max(id_factura) as id_factura from eer_facturas";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    ultimaFactura = Convert.ToInt32(r["id_factura"]);
                }
                db.CloseConnection();
                return ultimaFactura;
            }
            catch(Exception e)
            {
                return ultimaFactura;
            }
        }


        private double CoeficienteMes(DateTime fd, DateTime fh)
        {
            // Devuelve el coeficiente correspondientes al numero de dias a facturar en un mes.            

            double diasFactura = 0;
            double diasDelMes = 0;
            diasFactura = Convert.ToDouble((fh - fd).TotalDays + 1);
            diasDelMes = Convert.ToDouble(DateTime.DaysInMonth(fh.Year, fh.Month));            
            return Math.Round((diasFactura / diasDelMes),2); 

        }

    }
}

