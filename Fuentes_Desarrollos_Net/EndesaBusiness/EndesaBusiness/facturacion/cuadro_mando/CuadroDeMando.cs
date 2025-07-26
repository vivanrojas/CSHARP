using EndesaBusiness.contratacion.eexxi;
using EndesaBusiness.servidores;
using EndesaEntity.facturacion.cuadroDeMando;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;

namespace EndesaBusiness.facturacion.cuadro_mando
{
    public class CuadroDeMando
    {
        logs.Log ficheroLog;
        utilidades.Param param;
        Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> dic_electricidad_ES;
        Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> dic_electricidad_PT;
        Dictionary<int, EndesaEntity.facturacion.cuadroDeMando.InformeGas> dic_gas;

        DateTime fd = new DateTime();
        DateTime fh = new DateTime();
        DateTime fd_mesTrabajo = new DateTime();
        DateTime fh_mesTrabajo = new DateTime();
        EndesaBusiness.utilidades.Fechas utilFechas;

        EndesaBusiness.facturacion.cuadro_mando.Resumen resumen;
                
        EndesaBusiness.medida.Pendiente pendiente;
        EndesaBusiness.facturacion.redshift.Pendiente pendiente_sap;

        //EndesaBusiness.facturacion.cuadro_mando.Scoring scoring;
        EndesaBusiness.facturacion.cuadro_mando.Riesgo riesgo;
        EndesaBusiness.cartera.Cartera_SalesForce cartera;
        EndesaBusiness.facturacion.Anomalias anomalias;
        EndesaBusiness.facturacion.Autoconsumo autoconsumo;
        EndesaBusiness.facturacion.cuadro_mando.TopPortugal_Electricidad topPortugal;

        EndesaBusiness.facturacion.cuadro_mando.NRIs nris;
        
        FileInfo file;

        // Para el control de la ejecucion
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;



        EndesaBusiness.sigame.SIGAME sigame;
        bool mostrar_resumen = true;
        bool mostrar_electricidad = false;
        bool mostrar_gas = true;
        bool mostrar_precio_fijo = false;

        

        public CuadroDeMando()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Informe_CuadroDeMando");
            param = new utilidades.Param("ccmm_param", servidores.MySQLDB.Esquemas.FAC);
            utilFechas = new EndesaBusiness.utilidades.Fechas();
            ss_pp = new utilidades.Seguimiento_Procesos();
            pendiente = new medida.Pendiente();
            pendiente_sap = new redshift.Pendiente();
        }

        public void Proceso(DateTime fecha, bool automatico)
        {

            //bool existe_pendiente = false;
            
            if (!automatico || ss_pp.GetFecha_FinProceso("Facturación", "Cuadro Mando Facturación", "Cuadro Mando Facturación").Date < DateTime.Now.Date)
            {
                if (automatico)
                    ss_pp.Update_Fecha_Inicio("Facturación", "Cuadro Mando Facturación", "Cuadro Mando Facturación");               


                string[] listaArchivos = System.IO.Directory.GetFiles(param.GetValue("SalidaInforme"),
                    param.GetValue("NombreFicheroInforme") + "*.xlsx");

                for (int i = 0; i < listaArchivos.Length; i++)
                {
                    file = new FileInfo(listaArchivos[i]);
                    file.Delete();
                }


                if (fecha == DateTime.MinValue)
                {
                    fecha = utilFechas.UltimoDiaHabil();
                    fd_mesTrabajo = utilFechas.FD_MesTrabajo();
                    fh_mesTrabajo = utilFechas.FH_MesTrabajo();

                    fd = new DateTime(fecha.Year, fecha.Month, 1);
                    fh = utilFechas.UltimoDiaHabil();
                }
                else
                {
                    fd = new DateTime(fecha.Year, fecha.Month, 1);
                    fh = fecha;
                    fd_mesTrabajo = new DateTime(fecha.AddMonths(-1).Year, fecha.AddMonths(-1).Month, 1);
                    fh_mesTrabajo = new DateTime(fecha.AddMonths(-1).Year, fecha.AddMonths(-1).Month,
                        DateTime.DaysInMonth(fecha.AddMonths(-1).Year, fecha.AddMonths(-1).Month));
                }

                //existe_pendiente = pendiente.ExisteFechaInformePendiente(fecha);

                //if (existe_pendiente)

                // Si ha finalizado hoy el informe pendiente SAP lanzamos Cuadro de mando porque ya está disponible los datos del pendiente SAP                    
                if (ss_pp.GetFecha_FinProceso("Facturación", "Informe Pendiente BI", "Informe Pendiente BI").Date.ToShortDateString() == DateTime.Now.Date.ToShortDateString())
                {

                    if (mostrar_resumen)
                    {
                        resumen = new Resumen(Convert.ToInt32(fh.ToString("yyyyMM")));
                    }


                    if (mostrar_electricidad)
                    {
                        
                        Console.WriteLine("Cargando Pendiente Web");

                        //GUS 13-09-2024 - Tras la migración ya no se actualizará más el pendiente SCE por lo que este objeto no tendrá info
                        if (fecha.Date == utilFechas.UltimoDiaHabil().Date)
                        {
                            pendiente.CargaPendiente();
                        }
                        else
                        {
                            pendiente.CargaPendiente(fecha);
                        }


                        //scoring = new Scoring();
                        riesgo = new Riesgo();

                        //anomalias = new EndesaBusiness.facturacion.Anomalias();

                        dic_electricidad_ES = Carga_PS_AT(fd_mesTrabajo, fh_mesTrabajo);
                        //dic_electricidad_ES = Carga_Inventario_Rosetta(fd_mesTrabajo, fh_mesTrabajo);

                        cartera = new cartera.Cartera_SalesForce(dic_electricidad_ES.Select(z => z.Value.nif).ToList());

                        dic_electricidad_ES = CargaAlarmas(dic_electricidad_ES);

                        dic_electricidad_ES = CargaAgora(dic_electricidad_ES);

                        //dic_electricidad_ES =
                        //    CargaComplementosContrato(dic_electricidad_ES, fd_mesTrabajo, fh_mesTrabajo, "A01");

                        dic_electricidad_ES =
                            CargaComplementosContrato(dic_electricidad_ES, fd_mesTrabajo, fh_mesTrabajo, "L39");

                        dic_electricidad_ES = Carga_TAM(dic_electricidad_ES);


                        dic_electricidad_ES = Carga_TAM_Consumo(dic_electricidad_ES);

                        Console.WriteLine("Cargando Autoconsumo");
                        autoconsumo = new Autoconsumo();

                        topPortugal = new TopPortugal_Electricidad();

                        dic_electricidad_PT = topPortugal.dic;

                        dic_electricidad_PT = Carga_TopPortugal(dic_electricidad_PT);

                        nris = new NRIs();
                        nris.Copia_NRI();
                        nris.CargaMatrizNRI();
                    }

                    if (mostrar_gas)
                    {
                        // Cargamos el inventario de GAS
                        sigame = new sigame.SIGAME(fd_mesTrabajo, fh_mesTrabajo, true);
                        Gas(fd_mesTrabajo, fh_mesTrabajo);
                    }

                    Console.WriteLine("Generando Excel");
                    InformeExcelGas(automatico);

                    if (automatico)
                        ss_pp.Update_Fecha_Fin("Facturación", "Cuadro Mando Facturación", "Cuadro Mando Facturación");
                }
                else
                {
                    ss_pp.Update_Comentario("Facturación", "Cuadro Mando Facturación", "Cuadro Mando Facturación",
                        "No se ha generado aún el informe pendiente SAP.");
                    //EnvioCorreoNoHayPendiente();
                }
            }
        }

        private Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> Carga_TopPortugal(
            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> informe)
        {
            try
            {
                Console.WriteLine("Cargando TAM");
                EndesaBusiness.facturacion.TamElectricidad tam =
                    new TamElectricidad(informe.Select(z => z.Value.cups20).ToList());


                foreach (KeyValuePair<string, EndesaEntity.facturacion.cuadroDeMando.Informe> p in informe)
                {
                    p.Value.tam = tam.GetTamCups20(p.Value.cups20);

                    pendiente.GetCups13(p.Value.cups13);
                    if (pendiente.existe)
                    {
                        if(pendiente.aaaammPdte <= Convert.ToInt32(fh_mesTrabajo.ToString("yyyyMM")))
                        {
                            p.Value.estado_LTP = pendiente.estado;
                            p.Value.mes = pendiente.aaaammPdte;
                        }
                        else
                            p.Value.estado_LTP = "FACTURADO";
                    }
                    else
                        p.Value.estado_LTP = "FACTURADO";
                }


                return informe;

            }catch(Exception ex)
            {
                ficheroLog.AddError("CuadroDeMando.Carga_TopPortugal: " + ex.Message);
                return null;
            }
        }




        private Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> Carga_PS_AT(DateTime fd, DateTime fh)
        {

            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> d =
                new Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe>();
            try
            {
                #region PS_AT
                Console.WriteLine("Cargando PS_AT");
                EndesaBusiness.contratacion.PS_AT psat = new contratacion.PS_AT();
                

                foreach (KeyValuePair<string, EndesaEntity.contratacion.PS_AT_Tabla> p in psat.dic)
                {
                    EndesaEntity.facturacion.cuadroDeMando.Informe c = new EndesaEntity.facturacion.cuadroDeMando.Informe();

                                        

                    switch (p.Value.empresa)
                    {
                        case "EEXXI":
                            c.empresa = "EEXXI";
                            break;
                        case "EE":
                            c.empresa = "MT-España";
                            break;
                    }


                    c.origen_sistemas = p.Value.migrado_sap ? "SAP" : "SCE";
                    c.nif = p.Value.cif;
                    c.cliente = p.Value.nombre_cliente;
                    c.cups13 = p.Value.cups13;
                    c.cups20 = p.Value.cups20;
                    c.estado_contrato = p.Value.estado_contrato_descripcion;
                    c.provincia = p.Value.provincia;
                    c.tipo = "";
                    c.contratoPS = p.Value.contrext;
                    if (p.Value.fecha_prevista_baja > DateTime.MinValue)
                        c.fecha_prevista_baja = p.Value.fecha_prevista_baja;

                    if (p.Value.fecha_baja_contrato > DateTime.MinValue)
                        c.fecha_baja_contrato = p.Value.fecha_baja_contrato;

                    pendiente.GetCups13(c.cups13);
                    if (!pendiente.existe)
                        pendiente_sap.GetEstados(c.cups20);

                    if (pendiente.existe && pendiente.aaaammPdte <= Convert.ToInt32(fd.ToString("yyyyMM")))
                    {
                        c.estado_LTP = pendiente.estado;
                        c.mes = pendiente.aaaammPdte;
                    }else if(pendiente_sap.existe && pendiente_sap.aaaammPdte <= Convert.ToInt32(fd.ToString("yyyyMM")))
                    {
                        c.estado_LTP = pendiente_sap.descripcion_subestado;
                        c.mes = pendiente_sap.aaaammPdte;
                    }
                    else                     
                        c.estado_LTP = "FACTURADO";
                    
                        

                    //c.scoring = scoring.Prioridad(c.nif);
                    c.riesgo = riesgo.Es_Riesgo(c.nif);


                    EndesaEntity.facturacion.cuadroDeMando.Informe o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);
                    
                }

                #endregion

                #region Sofisticados Manuales (cm_sofisticados)
                Console.WriteLine("Cargando Sofisticados Manuales");

                Dictionary<string, EndesaEntity.facturacion.AgoraManual> agoraManual =
                    CargaAgoraManual(fd, fh);

                foreach(KeyValuePair<string, EndesaEntity.facturacion.AgoraManual> p in agoraManual)
                {
                    EndesaEntity.facturacion.cuadroDeMando.Informe o;
                    if (!d.TryGetValue(p.Value.cups20, out o))
                    {
                        EndesaEntity.facturacion.cuadroDeMando.Informe c = 
                            new EndesaEntity.facturacion.cuadroDeMando.Informe();

                        c.origen_sistemas = p.Value.migrado_sap ? "SAP" : "SCE";
                        c.empresa = p.Value.empresa;
                        c.nif = p.Value.nif;
                        c.cliente = p.Value.cliente;
                        c.cups13 = p.Value.cups13.ToUpper();
                        c.cups20 = p.Value.cups20.ToUpper();
                        c.estado_contrato = "EN VIGOR EN SOFISTICADOS";
                        c.tipo = "AGORA MANUAL";


                        pendiente.GetCups13(c.cups13);
                        if (pendiente.existe && pendiente.aaaammPdte <= Convert.ToInt32(fd.ToString("yyyyMM")))
                        {
                            c.estado_LTP = pendiente.estado;
                            c.mes = pendiente.aaaammPdte;
                        }
                        else
                            c.estado_LTP = "FACTURADO";

                        //c.scoring = scoring.Prioridad(c.nif);
                        c.riesgo = riesgo.Es_Riesgo(c.nif);

                        d.Add(c.cups20, c);
                    }else
                        o.tipo = "AGORA MANUAL";
                }

                #endregion

                return d;
            }
            catch(Exception ex)
            {
                ficheroLog.AddError("CuadroDeMando.Carga_PS_AT: " + ex.Message);
                return null;
            }

            

        }

        private Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> Carga_Inventario_Rosetta(DateTime fd, DateTime fh)
        {

            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> d =
                new Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe>();
            try
            {
                #region PS_AT
                Console.WriteLine("Cargando Inventario Rosetta");
                List<string> lista_cups_20 = new List<string>();
                EndesaBusiness.contratacion.Redshift.Inventario inventario_rosetta =
                    new contratacion.Redshift.Inventario(lista_cups_20);
                

                foreach (KeyValuePair<string, EndesaEntity.contratacion.Inventario> p in inventario_rosetta.dic)
                {
                    EndesaEntity.facturacion.cuadroDeMando.Informe c = new EndesaEntity.facturacion.cuadroDeMando.Informe();
                    c.cd_pais = p.Value.cd_pais;
                    c.empresa = p.Value.cd_empr;
                    c.nif = p.Value.cd_nif_cif_cli;
                    c.cliente = p.Value.tx_apell_cli;
                    c.cups13 = p.Value.cd_cups.ToUpper();
                    c.cups20 = p.Value.cups20.ToUpper();
                    c.estado_contrato = p.Value.de_estado_crto;
                    c.provincia = p.Value.de_prov;
                    c.tipo = "";
                    c.contratoPS = p.Value.cd_crto_comercial;
                    c.origen_sistemas = p.Value.lg_migrado_sap ? "SAP" : "SCE";

                   

                    if (p.Value.lg_migrado_sap)
                    {
                        // Pendiente SAP
                        c.estado_LTP = "FALTA INCLUIR PDTE SAP";
                        c.mes = Convert.ToInt32(fd.ToString("yyyyMM"));

                    }
                    else
                    {
                        c.cups13 = p.Value.cd_cups;
                        pendiente.GetCups13(c.cups13);
                        if (pendiente.existe && pendiente.aaaammPdte <= Convert.ToInt32(fd.ToString("yyyyMM")))
                        {
                            c.estado_LTP = pendiente.estado;
                            c.mes = pendiente.aaaammPdte;
                        }
                        else
                            c.estado_LTP = "FACTURADO";
                    }

                    

                    //c.scoring = scoring.Prioridad(c.nif);
                    c.riesgo = riesgo.Es_Riesgo(c.nif);


                    EndesaEntity.facturacion.cuadroDeMando.Informe o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);

                }

                #endregion

                #region Sofisticados Manuales (cm_sofisticados)
                Console.WriteLine("Cargando Sofisticados Manuales");

                Dictionary<string, EndesaEntity.facturacion.AgoraManual> agoraManual =
                    CargaAgoraManual(fd, fh);

                foreach (KeyValuePair<string, EndesaEntity.facturacion.AgoraManual> p in agoraManual)
                {
                    EndesaEntity.facturacion.cuadroDeMando.Informe o;
                    if (!d.TryGetValue(p.Value.cups20, out o))
                    {
                        EndesaEntity.facturacion.cuadroDeMando.Informe c =
                            new EndesaEntity.facturacion.cuadroDeMando.Informe();

                        c.origen_sistemas = "SCE";
                        c.empresa = p.Value.empresa;
                        c.nif = p.Value.nif;
                        c.cliente = p.Value.cliente;
                        c.cups13 = p.Value.cups13.ToUpper();
                        c.cups20 = p.Value.cups20.ToUpper();
                        c.estado_contrato = "EN VIGOR EN SOFISTICADOS";
                        c.tipo = "AGORA MANUAL";

                        pendiente.GetCups13(c.cups13);
                        if (pendiente.existe && pendiente.aaaammPdte <= Convert.ToInt32(fd.ToString("yyyyMM")))
                        {
                            c.estado_LTP = pendiente.estado;
                            c.mes = pendiente.aaaammPdte;
                        }
                        else
                            c.estado_LTP = "FACTURADO";

                        //c.scoring = scoring.Prioridad(c.nif);
                        c.riesgo = riesgo.Es_Riesgo(c.nif);

                        d.Add(c.cups20, c);
                    }
                    else
                        o.tipo = "AGORA MANUAL";
                }

                #endregion

                return d;
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("CuadroDeMando.Inventario_Rosetta: " + ex.Message);
                return null;
            }



        }

        private Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> Carga_TAM(
            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> informe)
        {

           
            try
            {
                Console.WriteLine("Cargando TAM");
                EndesaBusiness.facturacion.TamElectricidad tam =
                    new TamElectricidad(informe.Select(z => z.Value.cups20).ToList());



                foreach (KeyValuePair<string, EndesaEntity.facturacion.cuadroDeMando.Informe> p in informe)
                {
                    p.Value.tam = tam.GetTamCups20(p.Value.cups20);
                    if (p.Value.tam > Convert.ToDouble(param.GetValue("TOP1_TAM")))
                        p.Value.top = "TOP - URIARTE";
                    else if (p.Value.tam < Convert.ToDouble(param.GetValue("TOP1_TAM")) &&
                             p.Value.tam > Convert.ToDouble(param.GetValue("TOP2_TAM")))
                        p.Value.top = "TOP - 100";
                    else
                        p.Value.top = "Otros";

                }

                return informe;
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("Carga_TAM: " + ex.Message);
                return null;
            }



        }


        private Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> Carga_TAM_Consumo(
            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> informe)
        {


            try
            {
                Console.WriteLine("Cargando TAM Consumos");
                EndesaBusiness.facturacion.TAM_Consumo tam_consumos =
                    new TAM_Consumo();
                tam_consumos.Carga();



                foreach (KeyValuePair<string, EndesaEntity.facturacion.cuadroDeMando.Informe> p in informe)
                {

                    p.Value.consumo_medio = tam_consumos.GetConsumo(p.Value.cups20);                   

                }

                return informe;
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("Carga_TAM: " + ex.Message);
                return null;
            }



        }

        private Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> CargaAlarmas(
            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> informe)
        {


            try
            {
                Console.WriteLine("Cargando Alarmas");
                EndesaBusiness.facturacion.Alarmas alarmas =
                    new facturacion.Alarmas();

                alarmas.CargaSAP();
                alarmas.Carga();
                

                foreach (KeyValuePair<string, EndesaEntity.facturacion.cuadroDeMando.Informe> p in informe)
                {
                    p.Value.alarma = alarmas.GetAlarma(p.Value.cups20);
                }

                return informe;
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("CuadroDeMando.CargaAlarmas: " + ex.Message);
                return null;
            }



        }

        private Dictionary<string, EndesaEntity.facturacion.AgoraManual> CargaAgoraManual(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, EndesaEntity.facturacion.AgoraManual> d
                = new Dictionary<string, EndesaEntity.facturacion.AgoraManual>();

            try
            {
                //Buscamos AgoraManual en PT
                strSql = "Select s.Empresa, s.NIF, s.CCOUNIPS, s.CUPS20, s.DAPERSOC, pt.lg_migrado_sap"
                    + " from fact.cm_sofisticados s"
                    + " INNER JOIN cont.t_ed_h_ps_pt pt ON"
                    + " pt.cups20 = s.CUPS20"
                    + " where (s.FD <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " s.FH >= '" + fh.ToString("yyyy-MM-dd") + "' or s.FH is null)";
                    
                db = new MySQLDB(servidores.MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.AgoraManual c = new EndesaEntity.facturacion.AgoraManual();
                    c.empresa = r["Empresa"].ToString();
                    c.nif = r["NIF"].ToString();
                    c.cups13 = r["CCOUNIPS"].ToString();
                    c.cups20 = r["CUPS20"].ToString();
                    c.cliente = r["DAPERSOC"].ToString();

                    if (r["lg_migrado_sap"] != System.DBNull.Value)
                        c.migrado_sap = r["lg_migrado_sap"].ToString() == "S";

                    EndesaEntity.facturacion.AgoraManual o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);

                }
                db.CloseConnection();

                //Buscamos AgoraManual en ES
                strSql = "Select s.Empresa, s.NIF, s.CCOUNIPS, s.CUPS20, s.DAPERSOC, ps.lg_migrado_sap"
                    + " from fact.cm_sofisticados s"
                    + " INNER JOIN cont.t_ed_h_ps ps ON"
                    + " ps.cups20 = s.CUPS20"
                    + " where (s.FD <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " s.FH >= '" + fh.ToString("yyyy-MM-dd") + "' or s.FH is null)";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.AgoraManual c = new EndesaEntity.facturacion.AgoraManual();
                    c.empresa = r["Empresa"].ToString();
                    c.nif = r["NIF"].ToString();
                    c.cups13 = r["CCOUNIPS"].ToString();
                    c.cups20 = r["CUPS20"].ToString();
                    c.cliente = r["DAPERSOC"].ToString();

                    if (r["lg_migrado_sap"] != System.DBNull.Value)
                        c.migrado_sap = r["lg_migrado_sap"].ToString() == "S";

                    EndesaEntity.facturacion.AgoraManual o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);

                }
                db.CloseConnection();

                return d;
            }
            catch(Exception e)
            {
                return null;
            }
        }

        private Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> CargaAgora(
            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> informe)
        {
            Sofisticados sof = new Sofisticados();
            sof.Contruye_Sofisticados();
            bool tieneComplemento = false;

            foreach (KeyValuePair<string, EndesaEntity.facturacion.cuadroDeMando.Informe> p in informe)
            {
                tieneComplemento = sof.EsSofisticado(p.Value.cups13);

                if (tieneComplemento && p.Value.alarma == "")                    
                    p.Value.tipo = p.Value.origen_sistemas == "SAP" ? "AGORA SAP" : "AGORA SCE";
                else if (tieneComplemento)
                    p.Value.tipo = "AGORA MANUAL";
                else if (sof.EsPrecioFijo(p.Value.cups13, p.Value.mes))
                {                    
                    p.Value.tipo = "Precio Fijo";
                }
                //A petición de Ignacio Villar forzamos el punto Forces Electriques d Andorra como Agora 13/02/2024. Al final comentamos 
                //porque descubrimos una tabla (cm_sofisticados) en donde se gestionan los puntos AGORA MANUAL

                //else if (String.Equals(p.Value.cups13,"XXX0000000015"))
                //{
                //    p.Value.tipo = "AGORA MANUAL";
                //}
                    
                else if (p.Value.tipo == "")
                    p.Value.tipo = "OTROS";
                
            }
            return informe;
        }

        private Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> CargaComplementosContrato(
            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> informe, 
            DateTime fd, DateTime fh, string complemento)
        {
            bool tieneComplemento = false;

            try
            {
                Console.WriteLine("Cargando Complementos de Contrato (contratosPS " + complemento + ")");
                EndesaBusiness.contratacion.ComplementosContrato contratosPS =
                    new EndesaBusiness.contratacion.ComplementosContrato(fd, fh, complemento);



                foreach (KeyValuePair<string, EndesaEntity.facturacion.cuadroDeMando.Informe> p in informe)
                {

                    tieneComplemento = contratosPS.TieneComplemento(p.Value.cups20);

                    

                    if (tieneComplemento && complemento == "L39")
                    {
                        p.Value.tipo_CM = "PASSTHROUGH";
                    }





                }
                return informe;
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("CuadroDeMando.CargaComplementosContrato: " + ex.Message);
                return null;
            }



        }

        


        private List<EndesaEntity.facturacion.cuadroDeMando.Informe> TopUriarte()
        {
            List<EndesaEntity.facturacion.cuadroDeMando.Informe> lista
                = new List<EndesaEntity.facturacion.cuadroDeMando.Informe>();

            Console.WriteLine("Seleccionando Detalle Top Uriarte");
            foreach(KeyValuePair<string, EndesaEntity.facturacion.cuadroDeMando.Informe> p in dic_electricidad_ES)
            {
                //if(p.Value.tam > Convert.ToDouble(param.GetValue("TOP1_TAM")) && p.Value.empresa == "MT-España")
                //    lista.Add(p.Value);

                if ((p.Value.tam > Convert.ToDouble(param.GetValue("TOP1_TAM")) && p.Value.cd_pais == "ESPAÑA") ||
                    (p.Value.tam > Convert.ToDouble(param.GetValue("TOP1_TAM")) && p.Value.empresa == "MT-España"))
                    lista.Add(p.Value);
            }

            return lista;
        }

        private List<EndesaEntity.facturacion.cuadroDeMando.Informe> Top100()
        {
            List<EndesaEntity.facturacion.cuadroDeMando.Informe> lista
                = new List<EndesaEntity.facturacion.cuadroDeMando.Informe>();

            Console.WriteLine("Seleccionando Detalle Top 100");
            foreach (KeyValuePair<string, EndesaEntity.facturacion.cuadroDeMando.Informe> p in dic_electricidad_ES)
            {
                if (p.Value.tam < Convert.ToDouble(param.GetValue("TOP1_TAM")) &&
                    p.Value.tam > Convert.ToDouble(param.GetValue("TOP2_TAM")))
                    lista.Add(p.Value);
            }

            return lista;
        }

        private List<EndesaEntity.facturacion.cuadroDeMando.Informe> Sofisticados()
        {

            EndesaBusiness.facturacion.Anomalias anomalias = new Anomalias();

            List<EndesaEntity.facturacion.cuadroDeMando.Informe> lista
                = new List<EndesaEntity.facturacion.cuadroDeMando.Informe>();

            Console.WriteLine("Seleccionando Detalle Sofisticados");
            foreach (KeyValuePair<string, EndesaEntity.facturacion.cuadroDeMando.Informe> p in dic_electricidad_ES)
            {

                if (p.Value.tipo == "AGORA MANUAL" || p.Value.tipo == "AGORA SCE" || p.Value.tipo == "AGORA SAP")
                {
                    if (cartera.ExisteCartera(p.Value.nif))
                    {
                        p.Value.gestor = cartera.gestor;
                        p.Value.posicion_gestor = cartera.posicion;
                    }

                    anomalias.GetAnomalia(p.Value.cups20);
                    if (anomalias.existe)
                        p.Value.anomalia = anomalias.anomalia;

                    lista.Add(p.Value);
                }


                if (p.Value.tipo == "Precio Fijo" && p.Value.estado_LTP != "FACTURADO")
                {
                    if (cartera.ExisteCartera(p.Value.nif))
                    {
                        p.Value.gestor = cartera.gestor;
                        p.Value.posicion_gestor = cartera.posicion;
                    }

                    anomalias.GetAnomalia(p.Value.cups20);
                    if (anomalias.existe)
                        p.Value.anomalia = anomalias.anomalia;

                    lista.Add(p.Value);
                }
               


                




            }

            return lista;
        }

        private List<EndesaEntity.facturacion.cuadroDeMando.Informe> PrecioFijo()
        {

            EndesaBusiness.facturacion.Anomalias anomalias = new Anomalias();

            List<EndesaEntity.facturacion.cuadroDeMando.Informe> lista
                = new List<EndesaEntity.facturacion.cuadroDeMando.Informe>();

            Console.WriteLine("Seleccionando Detalle Precio Fijo");
            foreach (KeyValuePair<string, EndesaEntity.facturacion.cuadroDeMando.Informe> p in dic_electricidad_ES)
            {

                if (p.Value.tipo == "Precio Fijo" && p.Value.estado_LTP != "FACTURADO")
                {
                    if (cartera.ExisteCartera(p.Value.nif))
                    {
                        p.Value.gestor = cartera.gestor;
                        p.Value.posicion_gestor = cartera.posicion;
                    }

                    anomalias.GetAnomalia(p.Value.cups20);
                    if (anomalias.existe)
                        p.Value.anomalia = anomalias.anomalia;

                    lista.Add(p.Value);
                }


            }

            return lista;
        }
        private List<EndesaEntity.facturacion.cuadroDeMando.Informe> Passthrough()
        {
            List<EndesaEntity.facturacion.cuadroDeMando.Informe> lista
                = new List<EndesaEntity.facturacion.cuadroDeMando.Informe>();

            Console.WriteLine("Seleccionando Detalle Passthrough");
            foreach (KeyValuePair<string, EndesaEntity.facturacion.cuadroDeMando.Informe> p in dic_electricidad_ES)
            {
                if (p.Value.tipo_CM == "PASSTHROUGH")
                {
                    lista.Add(p.Value);
                }

            }

            return lista;
        }

        private List<EndesaEntity.facturacion.cuadroDeMando.Informe> XXI()
        {
            List<EndesaEntity.facturacion.cuadroDeMando.Informe> lista
                = new List<EndesaEntity.facturacion.cuadroDeMando.Informe>();


            double tam_top_xxi = Convert.ToDouble(param.GetValue("tam_top_xxi"));

            Console.WriteLine("Seleccionando Detalle XXI");
            foreach (KeyValuePair<string, EndesaEntity.facturacion.cuadroDeMando.Informe> p in dic_electricidad_ES)
            {
                if ((p.Value.empresa == "ENDESA ENERGÍA XXI" && p.Value.tam >= tam_top_xxi) ||
                    (p.Value.empresa == "EEXXI" && p.Value.tam >= tam_top_xxi))
                {
                    lista.Add(p.Value);
                }

            }

            return lista;
        }

        private void Gas(DateTime fd, DateTime fh)
        {
            EndesaBusiness.facturacion.FacturasGas facturasGAS = new FacturasGas();
            EndesaBusiness.facturacion.TamGas tamGas;
            EndesaBusiness.medida.GAS_MedidasFacturas medidasYfacturas;


            try
            {

                medidasYfacturas = new medida.GAS_MedidasFacturas();

                dic_gas = new Dictionary<int, EndesaEntity.facturacion.cuadroDeMando.InformeGas>();
                facturasGAS.CargaFacturasGasCuadrodeMando(fd.AddMonths(-12), fh);

                tamGas = new TamGas();

                foreach (KeyValuePair<string, EndesaEntity.sigame.ContratoGas> p in sigame.dic)
                {

                    EndesaEntity.facturacion.cuadroDeMando.InformeGas c = new EndesaEntity.facturacion.cuadroDeMando.InformeGas();
                    c.id_PS = p.Value.id_ps;
                    c.cif = p.Value.nif;
                    c.nombre_punto_suminsitro = p.Value.descripcion_ps;
                    c.grupo = p.Value.tarifa;
                    c.fecha_inicio = p.Value.fecha_inicio;
                    c.fecha_fin = p.Value.fecha_fin;
                    c.pais = p.Value.pais;
                    c.fecha_informe = DateTime.Now.Date;


                    c.promedio_facturacion = tamGas.GetTam_ID_PS(c.id_PS);
                    medidasYfacturas.GetID_PS(c.id_PS);

                    if (medidasYfacturas.existe)
                    {
                        c.medida = medidasYfacturas.medida; // Fecha de la medida

                        c.comentario = medidasYfacturas.comentario;
                        if (medidasYfacturas.mes == Convert.ToInt32(fd.ToString("yyyyMM")))
                            c.pendiente = "Pend. Facturación";
                        else
                            c.pendiente = "Pdte. Medida";

                    }


                    if (p.Value.cupsree.Length == 20)
                    {
                        c.es_cisterna = false;
                        c.cups = p.Value.cupsree;


                    }
                    else
                        c.es_cisterna = true;

                    // Buscamos datos de factura
                    facturasGAS.Get_ID_PS_PrimeraFactura(c.id_PS);
                    if (facturasGAS.existe)
                    {                        

                        if (facturasGAS.ultimo_mes_facturado == Convert.ToInt32(fd.ToString("yyyyMM")))
                        {
                            if (facturasGAS.cfactura == null)
                                c.pendiente = "Pdte"; //modificamos literal Pdte SCE --> Pdte
                            else
                                c.pendiente = "";
                        }

//                        if (facturasGAS.fecha_expedicion_factura > fh_mesTrabajo)
                            c.fecha_emision_sigame = facturasGAS.fecha_expedicion_factura;

                        c.ultimo_mes_facturado = facturasGAS.ultimo_mes_facturado;

                        if (facturasGAS.ffactura > DateTime.MinValue
                            //&& facturasGAS.ffactura >= fh.AddDays(1)
                            && (facturasGAS.cfactura != null && !facturasGAS.cfactura.Contains("S")))
                        {
                            c.ffactura = facturasGAS.ffactura;
                        }


                    }

                    EndesaEntity.facturacion.cuadroDeMando.InformeGas o;
                    if (!dic_gas.TryGetValue(c.id_PS, out o))
                    {
                        if((c.fecha_inicio <= DateTime.Now && c.fecha_fin >= DateTime.Now) ||
                            (c.fecha_fin <= DateTime.Now && (c.ultimo_mes_facturado > 0 &&
                            (c.ultimo_mes_facturado < Convert.ToInt32(c.fecha_fin.ToString("yyyyMM"))))))
                        dic_gas.Add(c.id_PS, c);
                    }
                        
                }





            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void InformeExcel(bool automatico)
        {
            int f = 1;
            int c = 1;

            DirectoryInfo dirSalida;
            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> dic_electricidad_ES =
                new Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe>();

            try
            {
                
                // Ruta de la plantilla 
                FileInfo plantillaExcel = new FileInfo(System.Environment.CurrentDirectory + param.GetValue("PlantillaInforme"));
                if(automatico)
                    dirSalida = new DirectoryInfo(param.GetValue("SalidaInforme"));
                else
                    dirSalida = new DirectoryInfo(@"C:\Temp\");


                FileInfo nombreSalidaExcel = new FileInfo(dirSalida.FullName
                    + param.GetValue("NombreFicheroInforme") 
                    + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);
                var workSheet = excelPackage.Workbook.Worksheets["Resumen"];
                var allCells = workSheet.Cells[1, 1, f, c];

                

                if (mostrar_electricidad)
                {
                    #region TOP Uriarte
                    workSheet = excelPackage.Workbook.Worksheets["TOP Uriarte"];

                    List<EndesaEntity.facturacion.cuadroDeMando.Informe> topUriarte = TopUriarte();

                    foreach (EndesaEntity.facturacion.cuadroDeMando.Informe p in topUriarte)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.origen_sistemas; c++;
                        workSheet.Cells[f, c].Value = p.nif; c++;
                        workSheet.Cells[f, c].Value = p.cliente; c++;
                        workSheet.Cells[f, c].Value = p.cups13; c++;
                        workSheet.Cells[f, c].Value = p.estado_contrato; c++;
                        workSheet.Cells[f, c].Value = p.provincia; c++;
                        if (p.mes > 0)
                            workSheet.Cells[f, c].Value = p.mes;
                        c++;
                        workSheet.Cells[f, c].Value = p.estado_LTP; c++;
                        workSheet.Cells[f, c].Value = p.tipo; c++;
                        workSheet.Cells[f, c].Value = p.tam;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        
                        if (p.riesgo)
                            workSheet.Cells[f, c].Value = "Sí";
                        else
                            workSheet.Cells[f, c].Value = "No";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    resumen.Save_Electricidad(topUriarte, Convert.ToInt32(fd.ToString("yyyyMM")),
                        "ES", "electricidad", "TOP URIARTE", fh.Day);

                    #endregion

                    #region TOP 100
                    workSheet = excelPackage.Workbook.Worksheets["TOP 100"];

                    List<EndesaEntity.facturacion.cuadroDeMando.Informe> top100 = Top100();

                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Informe p in top100)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.origen_sistemas; c++;
                        workSheet.Cells[f, c].Value = p.nif; c++;
                        workSheet.Cells[f, c].Value = p.cliente; c++;
                        workSheet.Cells[f, c].Value = p.cups13; c++;
                        workSheet.Cells[f, c].Value = p.estado_contrato; c++;
                        workSheet.Cells[f, c].Value = p.provincia; c++;
                        if (p.mes > 0)
                            workSheet.Cells[f, c].Value = p.mes;
                        c++;
                        workSheet.Cells[f, c].Value = p.estado_LTP; c++;
                        workSheet.Cells[f, c].Value = p.tipo; c++;
                        workSheet.Cells[f, c].Value = p.tam;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        
                        if(p.riesgo)
                            workSheet.Cells[f, c].Value = "Sí"; 
                        else
                            workSheet.Cells[f, c].Value = "No";

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    resumen.Save_Electricidad(top100, Convert.ToInt32(fd.ToString("yyyyMM")),
                        "ES", "electricidad", "TOP 100", fh.Day);
                    #endregion

                    #region Sofisticados
                    workSheet = excelPackage.Workbook.Worksheets["Sofisticados"];

                    List<EndesaEntity.facturacion.cuadroDeMando.Informe> sofisticados = Sofisticados();
                    //Creamos dos listas nuevas para sofisticados: SCE y SAP
                    List<EndesaEntity.facturacion.cuadroDeMando.Informe> sofisticados_SCE = new List<Informe>();
                    List<EndesaEntity.facturacion.cuadroDeMando.Informe> sofisticados_SAP = new List<Informe>();


                    EndesaBusiness.medida.CRRD curvaResumen =
                        new EndesaBusiness.medida.CRRD(sofisticados.Select(z => z.cups13).ToList(), "REGISTRADA");

                    EndesaBusiness.medida.Pendiente pendiente_Normal =
                        new medida.Pendiente(sofisticados.Select(z => z.cups13).ToList());

                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Informe p in sofisticados)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.origen_sistemas; c++;
                        //Comprobamos origen y añadimos a listas correspondientes
                        if(String.Equals(p.origen_sistemas,"SCE"))
                        {
                            sofisticados_SCE.Add(p);
                        }
                        else if (String.Equals(p.origen_sistemas, "SAP"))
                        {
                            sofisticados_SAP.Add(p);
                        }
                        workSheet.Cells[f, c].Value = p.nif; c++;
                        workSheet.Cells[f, c].Value = p.cliente; c++;
                        c++; // GRUPO
                        workSheet.Cells[f, c].Value = p.cups13; c++;
                        workSheet.Cells[f, c].Value = p.cups20; c++;
                        workSheet.Cells[f, c].Value = p.contratoPS; c++;
                        workSheet.Cells[f, c].Value = p.estado_contrato; c++;

                        if(p.fecha_baja_contrato > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.fecha_baja_contrato;
                            workSheet.Cells[f, c].Style.Numberformat.Format =
                                        DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }                        
                        c++;

                        workSheet.Cells[f, c].Value = p.provincia; c++;
                        if (p.mes > 0)
                            workSheet.Cells[f, c].Value = p.mes;
                        c++;
                        workSheet.Cells[f, c].Value = p.estado_LTP; c++;
                        workSheet.Cells[f, c].Value = p.top; c++;
                        workSheet.Cells[f, c].Value = p.tipo; c++;

                        workSheet.Cells[f, c].Value = p.tam;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                        
                        if (p.riesgo)
                            workSheet.Cells[f, c].Value = "Sí";
                        else
                            workSheet.Cells[f, c].Value = "No";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;


                        workSheet.Cells[f, c].Value = p.gestor; c++;
                        workSheet.Cells[f, c].Value = p.posicion_gestor; c++;

                        for (int i = 0; i <= 4; i++)
                        {
                            workSheet.Cells[f, c].Value = p.anomalia[i]; c++;
                        }
                        workSheet.Cells[f, c].Value = p.alarma; c++;

                        pendiente_Normal.GetCups13_Normal(p.cups13);
                        if (pendiente_Normal.existe)
                        {
                            EndesaEntity.medida.CurvaResumen ccrr =
                               curvaResumen.Primera_CR_DesdeFecha(p.cups13, pendiente_Normal.fh_desde);

                            if (pendiente_Normal.multimedida)
                            {
                                workSheet.Cells[f, c].Value = p.cups13; c++;
                                workSheet.Cells[f, c].Value = p.cups20; c++;
                                workSheet.Cells[f, c].Value = "S";
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;
                                if (ccrr != null)
                                {
                                    workSheet.Cells[f, c].Value = pendiente_Normal.fh_desde;
                                    workSheet.Cells[f, c].Style.Numberformat.Format =
                                        DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                                    workSheet.Cells[f, c].Value = ccrr.fh;
                                    workSheet.Cells[f, c].Style.Numberformat.Format =
                                        DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                                }                                   
                                
                            }                                 
                            else
                            {
                                workSheet.Cells[f, c].Value = pendiente_Normal.cups15; c++;
                                if (ccrr != null)
                                    workSheet.Cells[f, c].Value = ccrr.cups22;
                                c++;
                                workSheet.Cells[f, c].Value = "N"; 
                                workSheet.Cells[f, c].Style.HorizontalAlignment = 
                                    ExcelHorizontalAlignment.Center; c++;

                                if (ccrr != null)
                                {
                                    workSheet.Cells[f, c].Value = pendiente_Normal.fh_desde;
                                    workSheet.Cells[f, c].Style.Numberformat.Format =
                                        DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                                    workSheet.Cells[f, c].Value = ccrr.fh;
                                    workSheet.Cells[f, c].Style.Numberformat.Format =
                                        DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                                }
                            }

                                  
                                                      
                            
                           

                        }

                    }

                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();


                    resumen.Save_Electricidad(sofisticados, Convert.ToInt32(fd.ToString("yyyyMM")),
                       "ES", "electricidad", "SOFISTICADOS", fh.Day);
                    //Dividimos resumen 'ES-electricidad-SOFISTICADOS' por 'ES-electricidad-SOFISTICADOS SCE' y 'ES-electricidad-SOFISTICADOS SAP'
                    resumen.Save_Electricidad(sofisticados_SCE, Convert.ToInt32(fd.ToString("yyyyMM")),
                       "ES", "electricidad", "SOFISTICADOS SCE", fh.Day);
                    resumen.Save_Electricidad_SAP(sofisticados_SAP, Convert.ToInt32(fd.ToString("yyyyMM")),
                       "ES", "electricidad", "SOFISTICADOS SAP", fh.Day);


                    #endregion

                    #region Precio Fijo
                    if (mostrar_precio_fijo)
                    {
                        workSheet = excelPackage.Workbook.Worksheets["Precio Fijo"];

                        List<EndesaEntity.facturacion.cuadroDeMando.Informe> prefio_fijo = PrecioFijo();

                       

                        f = 1;
                        foreach (EndesaEntity.facturacion.cuadroDeMando.Informe p in prefio_fijo)
                        {
                            c = 1;
                            f++;
                            workSheet.Cells[f, c].Value = p.origen_sistemas; c++;
                            workSheet.Cells[f, c].Value = p.nif; c++;
                            workSheet.Cells[f, c].Value = p.cliente; c++;
                            c++; // GRUPO
                            workSheet.Cells[f, c].Value = p.cups13; c++;
                            workSheet.Cells[f, c].Value = p.cups20; c++;
                            workSheet.Cells[f, c].Value = p.contratoPS; c++;
                            workSheet.Cells[f, c].Value = p.estado_contrato; c++;
                            workSheet.Cells[f, c].Value = p.provincia; c++;
                            if (p.mes > 0)
                                workSheet.Cells[f, c].Value = p.mes;
                            c++;
                            workSheet.Cells[f, c].Value = p.estado_LTP; c++;
                            workSheet.Cells[f, c].Value = p.top; c++;
                            workSheet.Cells[f, c].Value = p.tipo; c++;

                            workSheet.Cells[f, c].Value = p.tam;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                            
                            if (p.riesgo)
                                workSheet.Cells[f, c].Value = "Sí";
                            else
                                workSheet.Cells[f, c].Value = "No";
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;


                            workSheet.Cells[f, c].Value = p.gestor; c++;
                            workSheet.Cells[f, c].Value = p.posicion_gestor; c++;

                            for (int i = 0; i <= 4; i++)
                            {
                                workSheet.Cells[f, c].Value = p.anomalia[i]; c++;
                            }
                            workSheet.Cells[f, c].Value = p.alarma; c++;

                            

                        }

                        allCells = workSheet.Cells[1, 1, f, c];
                        allCells.AutoFitColumns();

                        resumen.Save_Electricidad(prefio_fijo, Convert.ToInt32(fd.ToString("yyyyMM")),
                           "ES", "electricidad", "PRECIO FIJO", fh.Day);
                        #endregion

                    }

                    #region Pass Through España
                    //workSheet = excelPackage.Workbook.Worksheets["Pass Through España"];

                    //List<EndesaEntity.facturacion.cuadroDeMando.Informe> passthrough = Passthrough();

                    //f = 1;
                    //foreach (EndesaEntity.facturacion.cuadroDeMando.Informe p in passthrough)
                    //{
                    //    c = 1;
                    //    f++;
                    //    workSheet.Cells[f, c].Value = p.origen_sistemas; c++;
                    //    workSheet.Cells[f, c].Value = p.nif; c++;
                    //    workSheet.Cells[f, c].Value = p.cliente; c++;
                    //    workSheet.Cells[f, c].Value = p.cups13; c++;
                    //    workSheet.Cells[f, c].Value = p.estado_contrato; c++;
                    //    workSheet.Cells[f, c].Value = p.provincia; c++;
                    //    if (p.mes > 0)
                    //        workSheet.Cells[f, c].Value = p.mes;
                    //    c++;
                    //    workSheet.Cells[f, c].Value = p.estado_LTP; c++;
                    //    workSheet.Cells[f, c].Value = p.top; c++;
                    //    workSheet.Cells[f, c].Value = p.tam;
                    //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    //    //workSheet.Cells[f, c].Value = p.scoring;
                    //    if (p.riesgo)
                    //        workSheet.Cells[f, c].Value = "Sí";
                    //    else
                    //        workSheet.Cells[f, c].Value = "No";
                    //    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //}
                    //allCells = workSheet.Cells[1, 1, f, c];
                    //allCells.AutoFitColumns();

                    //resumen.Save_Electricidad(passthrough, Convert.ToInt32(fd.ToString("yyyyMM")),
                    //    "ES", "electricidad", "PASS THROUGH ESPAÑA", fh.Day);
                    
                    

                    #endregion

                    #region XXI
                    workSheet = excelPackage.Workbook.Worksheets["EXXI"];

                    List<EndesaEntity.facturacion.cuadroDeMando.Informe> xxi = XXI();

                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Informe p in xxi)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.origen_sistemas; c++;
                        workSheet.Cells[f, c].Value = p.nif; c++;
                        workSheet.Cells[f, c].Value = p.cliente; c++;
                        workSheet.Cells[f, c].Value = p.cups13; c++;
                        workSheet.Cells[f, c].Value = p.estado_contrato; c++;
                        workSheet.Cells[f, c].Value = p.provincia; c++;
                        if (p.mes > 0)
                            workSheet.Cells[f, c].Value = p.mes;
                        c++;
                        workSheet.Cells[f, c].Value = p.estado_LTP; c++;
                        workSheet.Cells[f, c].Value = p.top; c++;
                        workSheet.Cells[f, c].Value = p.tam;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        //workSheet.Cells[f, c].Value = p.scoring;
                        if (p.riesgo)
                            workSheet.Cells[f, c].Value = "Sí";
                        else
                            workSheet.Cells[f, c].Value = "No";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; 
                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    resumen.Save_Electricidad(xxi, Convert.ToInt32(fd.ToString("yyyyMM")),
                       "ES", "electricidad", "ENDESA ENERGÍA XXI", fh.Day);

                    #endregion

                    #region Autoconsumo
                    workSheet = excelPackage.Workbook.Worksheets["Autoconsumo"];

                    List<EndesaEntity.facturacion.cuadroDeMando.Autoconsumo> lista_autoconsumo =
                        autoconsumo.dic.Values.ToList();


                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Autoconsumo p in lista_autoconsumo)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.origen_sistemas; c++;
                        workSheet.Cells[f, c].Value = p.empresa; c++;
                        workSheet.Cells[f, c].Value = p.nif; c++;
                        workSheet.Cells[f, c].Value = p.cliente; c++;
                        workSheet.Cells[f, c].Value = p.cups22; c++;
                        workSheet.Cells[f, c].Value = p.tarifa; c++;
                        workSheet.Cells[f, c].Value = p.descripcion; c++;

                        
                        //if (riesgo.Es_Riesgo(p.nif))
                        //    workSheet.Cells[f, c].Value = "Sí";
                        //else
                        //    workSheet.Cells[f, c].Value = "No";
                        //workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();
                    
                    #endregion

                    #region Top Portugal
                    workSheet = excelPackage.Workbook.Worksheets["TOP Portugal"];

                    List<EndesaEntity.facturacion.cuadroDeMando.Informe> topPortugal =
                        dic_electricidad_PT.Values.OrderBy(z => z.estado_LTP).ToList();

                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Informe p in topPortugal)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.origen_sistemas; c++;
                        workSheet.Cells[f, c].Value = p.nif; c++;
                        workSheet.Cells[f, c].Value = p.cliente; c++;
                        workSheet.Cells[f, c].Value = p.cups20; c++;
                        if(p.mes > 0)
                            workSheet.Cells[f, c].Value = p.mes; 
                        c++;
                        workSheet.Cells[f, c].Value = p.estado_LTP; c++;
                        workSheet.Cells[f, c].Value = p.tam;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    resumen.Save_Electricidad(topPortugal, Convert.ToInt32(fd.ToString("yyyyMM")),
                       "PT", "electricidad", "TOP PORTUGAL", fh.Day);

                    #endregion

                    #region NRI´s
                    workSheet = excelPackage.Workbook.Worksheets["NRI"];
                    f = 1;

                    Console.WriteLine("Seleccionando NRI´s");

                    foreach (KeyValuePair<int, EndesaEntity.facturacion.cuadroDeMando.NRI> p in nris.dic)
                    {
                        c = 1;
                        f++;

                        workSheet.Cells[f, c].Value = p.Value.fecha_ultimo_estado;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                        workSheet.Cells[f, c].Value = p.Value.estado; c++;
                        workSheet.Cells[f, c].Value = p.Value.codigo_nri; c++;
                        workSheet.Cells[f, c].Value = p.Value.cliente; c++;
                        workSheet.Cells[f, c].Value = p.Value.nif; c++;
                        workSheet.Cells[f, c].Value = p.Value.plazo; c++;
                        workSheet.Cells[f, c].Value = p.Value.motivo_alta; c++;
                        workSheet.Cells[f, c].Value = p.Value.submotivo_alta; c++;
                        workSheet.Cells[f, c].Value = p.Value.linea_negocio; 

                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    #endregion
                }


                if (mostrar_gas)
                {
                    #region Grupos Especiales
                    Console.WriteLine("Grupos Especiales");
                    workSheet = excelPackage.Workbook.Worksheets["Grupos Especiales"];
                    
                    new List<EndesaEntity.facturacion.cuadroDeMando.InformeGas>();

                    GrupoCliente grupocliente = new GrupoCliente();

                    


                    List<EndesaEntity.facturacion.cuadroDeMando.InformeGas> grupos_especiales = Trata_Grupos_Especiales(grupocliente.dic);

                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.InformeGas p in grupos_especiales)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.id_PS; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.nombre_punto_suminsitro; c++;
                        workSheet.Cells[f, c].Value = p.cups; c++;
                        workSheet.Cells[f, c].Value = p.grupo; c++;
                        workSheet.Cells[f, c].Value = p.fecha_inicio;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        if (p.fecha_fin < new DateTime(4999, 12, 31))
                        {
                            workSheet.Cells[f, c].Value = p.fecha_fin;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.ultimo_mes_facturado; c++;

                        workSheet.Cells[f, c].Value = p.medida;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.fecha_emision_sigame;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;


                        workSheet.Cells[f, c].Value = p.fecha_informe;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        // DIF (F. INFORME - F. EMISIÓN)
                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_informe - p.fecha_emision_sigame).TotalDays);

                        }
                        c++;

                        if (p.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_emision_sigame - p.ffactura).TotalDays);
                            c++;

                        }
                        else
                        {
                            c++; // FFACTURA
                            c++; // DIF (FFACTURA - F.EMISIÓN)                        
                        }


                        workSheet.Cells[f, c].Value = p.pendiente; c++;
                        workSheet.Cells[f, c].Value = p.comentario; c++;
                        workSheet.Cells[f, c].Value = p.promedio_facturacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    if (mostrar_resumen)
                    {
                        resumen.Save_GAS(grupos_especiales, Convert.ToInt32(fd.ToString("yyyyMM")),
                            "ES", "gas", "Grupos Especiales", fh.Day);
                    }


                    #endregion

                    #region TOP gas España
                    Console.WriteLine("Hoja TOP gas España");
                    workSheet = excelPackage.Workbook.Worksheets["TOP gas España"];
                    List<EndesaEntity.facturacion.cuadroDeMando.InformeGas> top_gas_espana =
                        dic_gas.Values.Where(z => z.pais == "España" && 
                        z.es_cisterna == false && z.promedio_facturacion > 100000).ToList();


                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.InformeGas p in top_gas_espana)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.id_PS; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.nombre_punto_suminsitro; c++;                        
                        workSheet.Cells[f, c].Value = p.cups; c++;
                        workSheet.Cells[f, c].Value = p.grupo; c++;
                        workSheet.Cells[f, c].Value = p.fecha_inicio;                        
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        if(p.fecha_fin < new DateTime(4999, 12, 31))
                        {
                            workSheet.Cells[f, c].Value = p.fecha_fin;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.ultimo_mes_facturado; c++;
                        
                        workSheet.Cells[f, c].Value = p.medida;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                        
                        if(p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.fecha_emision_sigame;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; 
                        }
                        c++;


                        workSheet.Cells[f, c].Value = p.fecha_informe;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        // DIF (F. INFORME - F. EMISIÓN)
                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_informe - p.fecha_emision_sigame).TotalDays);
                            
                        }
                        c++; 

                        if(p.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_emision_sigame - p.ffactura).TotalDays);
                            c++;    

                        }
                        else
                        {
                            c++; // FFACTURA
                            c++; // DIF (FFACTURA - F.EMISIÓN)                        
                        }
                                                
                        
                        workSheet.Cells[f, c].Value = p.pendiente; c++;
                        workSheet.Cells[f, c].Value = p.comentario; c++;
                        workSheet.Cells[f, c].Value = p.promedio_facturacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    if (mostrar_resumen)
                    {
                        resumen.Save_GAS(top_gas_espana, Convert.ToInt32(fd.ToString("yyyyMM")),
                            "ES", "gas", "TOP", fh.Day);
                    }



                    #endregion

                    #region NO TOP gas España
                    Console.WriteLine("Hoja NO TOP gas España");
                    workSheet = excelPackage.Workbook.Worksheets["NO TOP gas España"];
                    List<EndesaEntity.facturacion.cuadroDeMando.InformeGas> no_top_gas_espana =
                        dic_gas.Values.Where(z => z.pais == "España" &&
                        z.es_cisterna == false && z.promedio_facturacion <= 100000).ToList();


                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.InformeGas p in no_top_gas_espana)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.id_PS; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.nombre_punto_suminsitro; c++;
                        workSheet.Cells[f, c].Value = p.cups; c++;
                        workSheet.Cells[f, c].Value = p.grupo; c++;
                        workSheet.Cells[f, c].Value = p.fecha_inicio;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        if (p.fecha_fin < new DateTime(4999, 12, 31))
                        {
                            workSheet.Cells[f, c].Value = p.fecha_fin;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.ultimo_mes_facturado; c++;

                        workSheet.Cells[f, c].Value = p.medida;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.fecha_emision_sigame;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;


                        workSheet.Cells[f, c].Value = p.fecha_informe;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        // DIF (F. INFORME - F. EMISIÓN)
                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_informe - p.fecha_emision_sigame).TotalDays);

                        }
                        c++;

                        if (p.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_emision_sigame - p.ffactura).TotalDays);
                            c++;

                        }
                        else
                        {
                            c++; // FFACTURA
                            c++; // DIF (FFACTURA - F.EMISIÓN)                        
                        }




                        workSheet.Cells[f, c].Value = p.pendiente; c++;
                        workSheet.Cells[f, c].Value = p.comentario; c++;
                        workSheet.Cells[f, c].Value = p.promedio_facturacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    if (mostrar_resumen)
                    {
                        resumen.Save_GAS(no_top_gas_espana, Convert.ToInt32(fd.ToString("yyyyMM")),
                            "ES", "gas", "NO TOP", fh.Day);
                    }


                    #endregion

                    #region Cisternas gas España
                    Console.WriteLine("cisternas gas España");
                    workSheet = excelPackage.Workbook.Worksheets["cisternas gas España"];
                    List<EndesaEntity.facturacion.cuadroDeMando.InformeGas> cisternas_gas_espana =
                        dic_gas.Values.Where(z => z.pais == "España" &&
                        z.es_cisterna == true).ToList();


                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.InformeGas p in cisternas_gas_espana)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.id_PS; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.nombre_punto_suminsitro; c++;
                        workSheet.Cells[f, c].Value = p.cups; c++;
                        workSheet.Cells[f, c].Value = p.grupo; c++;
                        workSheet.Cells[f, c].Value = p.fecha_inicio;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        if (p.fecha_fin < new DateTime(4999, 12, 31))
                        {
                            workSheet.Cells[f, c].Value = p.fecha_fin;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.ultimo_mes_facturado; c++;

                        workSheet.Cells[f, c].Value = p.medida;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.fecha_emision_sigame;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;


                        workSheet.Cells[f, c].Value = p.fecha_informe;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        // DIF (F. INFORME - F. EMISIÓN)
                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_informe - p.fecha_emision_sigame).TotalDays);

                        }
                        c++;

                        if (p.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_emision_sigame - p.ffactura).TotalDays);
                            c++;

                        }
                        else
                        {
                            c++; // FFACTURA
                            c++; // DIF (FFACTURA - F.EMISIÓN)                        
                        }




                        workSheet.Cells[f, c].Value = p.pendiente; c++;
                        workSheet.Cells[f, c].Value = p.comentario; c++;
                        workSheet.Cells[f, c].Value = p.promedio_facturacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    if (mostrar_resumen)
                    {
                        resumen.Save_GAS(cisternas_gas_espana, Convert.ToInt32(fd.ToString("yyyyMM")),
                            "ES", "gas", "CISTERNAS", fh.Day);
                    }


                    #endregion

                    #region Gas Portugal
                    Console.WriteLine("Hoja gas Portugal");
                    workSheet = excelPackage.Workbook.Worksheets["gas portugal"];
                    List<EndesaEntity.facturacion.cuadroDeMando.InformeGas> gasPortugal =
                        dic_gas.Values.Where(z => z.pais == "Portugal").ToList();


                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.InformeGas p in gasPortugal)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.id_PS; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.nombre_punto_suminsitro; c++;
                        workSheet.Cells[f, c].Value = p.cups; c++;                        
                        workSheet.Cells[f, c].Value = p.fecha_inicio;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        if (p.fecha_fin < new DateTime(4999, 12, 31))
                        {
                            workSheet.Cells[f, c].Value = p.fecha_fin;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.ultimo_mes_facturado; c++;

                        workSheet.Cells[f, c].Value = p.medida;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.fecha_emision_sigame;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;


                        workSheet.Cells[f, c].Value = p.fecha_informe;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        // DIF (F. INFORME - F. EMISIÓN)
                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_informe - p.fecha_emision_sigame).TotalDays);

                        }
                        c++;

                        if (p.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_emision_sigame - p.ffactura).TotalDays);
                            c++;

                        }
                        else
                        {
                            c++; // FFACTURA
                            c++; // DIF (FFACTURA - F.EMISIÓN)                        
                        }




                        workSheet.Cells[f, c].Value = p.pendiente; c++;
                        workSheet.Cells[f, c].Value = p.comentario; c++;
                        workSheet.Cells[f, c].Value = p.promedio_facturacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    if (mostrar_resumen)
                    {
                        resumen.Save_GAS(gasPortugal, Convert.ToInt32(fd.ToString("yyyyMM")),
                            "PT", "gas", "TOTAL", fh.Day);
                    }
                       

                    #endregion
                }

                if (mostrar_resumen)
                {
                    Console.WriteLine("Hoja Resumen");
                    workSheet = excelPackage.Workbook.Worksheets["Resumen"];

                    // Cabecera
                    workSheet.Cells[3, 1].Value = utilFechas.ConvierteMes_a_Letra(fd) 
                        + " " + fd.Year;


                    // Pintamos dias                   
                    f = 7;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }                        
                        

                    f = 20;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }

                    f = 33;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }

                    f = 46;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }

                    f = 55;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }

                    f = 70;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }

                    f = 85;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }

                    f = 91;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }

                    f = 97;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }

                    f = 105;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }

                    //f = 109;
                    //c = 3;
                    //for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    //{
                    //    workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    //}


                    List<EndesaEntity.facturacion.cuadroDeMando.Resumen> lista_resumen
                        = resumen.GetValor("ES", "electricidad", "TOP URIARTE");

                    f = 7;                    
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {
                        
                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {
                            
                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }

                    lista_resumen  = resumen.GetValor("ES", "electricidad", "TOP 100");

                    f = 20;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {

                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {

                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }

                    lista_resumen = resumen.GetValor("ES", "electricidad", "SOFISTICADOS SCE");

                    f = 33;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {

                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {

                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }

                    lista_resumen = resumen.GetValor("ES", "electricidad", "SOFISTICADOS SAP");

                    f = 46;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {

                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {

                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }

                    //lista_resumen = resumen.GetValor("ES", "electricidad", "PASS THROUGH ESPAÑA");

                    //f = 46;
                    //foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    //{

                    //    c = 3;
                    //    f++;
                    //    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    //    {

                    //        if (p.dias[d.Day] >= 0)
                    //            workSheet.Cells[f, c].Value = p.dias[d.Day];
                    //        c++;
                    //    }

                    //}

                    lista_resumen = resumen.GetValor("ES", "electricidad", "ENDESA ENERGÍA XXI");

                    f = 55;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {

                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {

                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }


                    lista_resumen = resumen.GetValor("PT", "electricidad", "TOP PORTUGAL");

                    f = 70;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {

                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {

                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }

                    #region GAS

                    lista_resumen = resumen.GetValor("ES", "gas", "TOP");

                    f = 85;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {

                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {

                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }

                    lista_resumen = resumen.GetValor("ES", "gas", "NO TOP");

                    f = 91;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {

                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {

                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }


                    lista_resumen = resumen.GetValor("ES", "gas", "CISTERNAS");

                    f = 97;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {

                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {

                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }


                    lista_resumen = resumen.GetValor("PT", "gas", "TOTAL");

                    f = 105;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {

                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {

                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }

                    #region NRI´s
                    f = 113;

                    if (nris != null)
                    {
                        for (int fil = 0; fil <= 3; fil++)
                        {
                            int columna = 0;
                            f++;
                            for (int col = 0; col <= 4; col++)
                            {
                                columna = columna + 3;
                                workSheet.Cells[f, columna].Value = nris.resumenNRI[fil, col];
                            }
                        }
                    }
                    #endregion

                    #endregion

                }

                excelPackage.SaveAs(nombreSalidaExcel);
                excelPackage = null;

                if (automatico)
                    EnvioCorreoInformeCuadroMando(nombreSalidaExcel.FullName);

                //ControlEstadoPuntoFEDA();
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("InformeExcel: " + ex.Message);
            }
        }

        private void InformeExcelGas(bool automatico)
        {
            int f = 1;
            int c = 1;

            DirectoryInfo dirSalida;
            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> dic_electricidad_ES =
                new Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe>();

            try
            {

                // Ruta de la plantilla 
                FileInfo plantillaExcel = new FileInfo(System.Environment.CurrentDirectory + param.GetValue("PlantillaInforme"));
                if (automatico)
                    dirSalida = new DirectoryInfo(param.GetValue("SalidaInforme"));
                else
                    dirSalida = new DirectoryInfo(@"C:\Temp\");


                FileInfo nombreSalidaExcel = new FileInfo(dirSalida.FullName
                    + param.GetValue("NombreFicheroInforme")
                    + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);
                var workSheet = excelPackage.Workbook.Worksheets["Resumen"];
                var allCells = workSheet.Cells[1, 1, f, c];


                if (mostrar_gas)
                {
                    #region Grupos Especiales
                    Console.WriteLine("Grupos Especiales");
                    workSheet = excelPackage.Workbook.Worksheets["Grupos Especiales"];

                    new List<EndesaEntity.facturacion.cuadroDeMando.InformeGas>();

                    GrupoCliente grupocliente = new GrupoCliente();

                    List<EndesaEntity.facturacion.cuadroDeMando.InformeGas> grupos_especiales = Trata_Grupos_Especiales(grupocliente.dic);

                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.InformeGas p in grupos_especiales)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.id_PS; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.nombre_punto_suminsitro; c++;
                        workSheet.Cells[f, c].Value = p.cups; c++;
                        workSheet.Cells[f, c].Value = p.grupo; c++;
                        workSheet.Cells[f, c].Value = p.fecha_inicio;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        if (p.fecha_fin < new DateTime(4999, 12, 31))
                        {
                            workSheet.Cells[f, c].Value = p.fecha_fin;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.ultimo_mes_facturado; c++;

                        workSheet.Cells[f, c].Value = p.medida;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.fecha_emision_sigame;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;


                        workSheet.Cells[f, c].Value = p.fecha_informe;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        // DIF (F. INFORME - F. EMISIÓN)
                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_informe - p.fecha_emision_sigame).TotalDays);

                        }
                        c++;

                        if (p.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_emision_sigame - p.ffactura).TotalDays);
                            c++;

                        }
                        else
                        {
                            c++; // FFACTURA
                            c++; // DIF (FFACTURA - F.EMISIÓN)                        
                        }


                        workSheet.Cells[f, c].Value = p.pendiente; c++;
                        workSheet.Cells[f, c].Value = p.comentario; c++;
                        workSheet.Cells[f, c].Value = p.promedio_facturacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    if (mostrar_resumen)
                    {
                        resumen.Save_GAS(grupos_especiales, Convert.ToInt32(fd.ToString("yyyyMM")),
                            "ES", "gas", "Grupos Especiales", fh.Day);
                    }


                    #endregion

                    #region TOP gas España
                    Console.WriteLine("Hoja TOP gas España");
                    workSheet = excelPackage.Workbook.Worksheets["TOP gas España"];
                    List<EndesaEntity.facturacion.cuadroDeMando.InformeGas> top_gas_espana =
                        dic_gas.Values.Where(z => z.pais == "España" &&
                        z.es_cisterna == false && z.promedio_facturacion > 100000).ToList();


                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.InformeGas p in top_gas_espana)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.id_PS; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.nombre_punto_suminsitro; c++;
                        workSheet.Cells[f, c].Value = p.cups; c++;
                        workSheet.Cells[f, c].Value = p.grupo; c++;
                        workSheet.Cells[f, c].Value = p.fecha_inicio;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        if (p.fecha_fin < new DateTime(4999, 12, 31))
                        {
                            workSheet.Cells[f, c].Value = p.fecha_fin;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.ultimo_mes_facturado; c++;

                        workSheet.Cells[f, c].Value = p.medida;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.fecha_emision_sigame;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;


                        workSheet.Cells[f, c].Value = p.fecha_informe;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        // DIF (F. INFORME - F. EMISIÓN)
                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_informe - p.fecha_emision_sigame).TotalDays);

                        }
                        c++;

                        if (p.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_emision_sigame - p.ffactura).TotalDays);
                            c++;

                        }
                        else
                        {
                            c++; // FFACTURA
                            c++; // DIF (FFACTURA - F.EMISIÓN)                        
                        }


                        workSheet.Cells[f, c].Value = p.pendiente; c++;
                        workSheet.Cells[f, c].Value = p.comentario; c++;
                        workSheet.Cells[f, c].Value = p.promedio_facturacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    if (mostrar_resumen)
                    {
                        resumen.Save_GAS(top_gas_espana, Convert.ToInt32(fd.ToString("yyyyMM")),
                            "ES", "gas", "TOP", fh.Day);
                    }



                    #endregion

                    #region NO TOP gas España
                    Console.WriteLine("Hoja NO TOP gas España");
                    workSheet = excelPackage.Workbook.Worksheets["NO TOP gas España"];
                    List<EndesaEntity.facturacion.cuadroDeMando.InformeGas> no_top_gas_espana =
                        dic_gas.Values.Where(z => z.pais == "España" &&
                        z.es_cisterna == false && z.promedio_facturacion <= 100000).ToList();


                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.InformeGas p in no_top_gas_espana)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.id_PS; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.nombre_punto_suminsitro; c++;
                        workSheet.Cells[f, c].Value = p.cups; c++;
                        workSheet.Cells[f, c].Value = p.grupo; c++;
                        workSheet.Cells[f, c].Value = p.fecha_inicio;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        if (p.fecha_fin < new DateTime(4999, 12, 31))
                        {
                            workSheet.Cells[f, c].Value = p.fecha_fin;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.ultimo_mes_facturado; c++;

                        workSheet.Cells[f, c].Value = p.medida;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.fecha_emision_sigame;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;


                        workSheet.Cells[f, c].Value = p.fecha_informe;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        // DIF (F. INFORME - F. EMISIÓN)
                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_informe - p.fecha_emision_sigame).TotalDays);

                        }
                        c++;

                        if (p.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_emision_sigame - p.ffactura).TotalDays);
                            c++;

                        }
                        else
                        {
                            c++; // FFACTURA
                            c++; // DIF (FFACTURA - F.EMISIÓN)                        
                        }




                        workSheet.Cells[f, c].Value = p.pendiente; c++;
                        workSheet.Cells[f, c].Value = p.comentario; c++;
                        workSheet.Cells[f, c].Value = p.promedio_facturacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    if (mostrar_resumen)
                    {
                        resumen.Save_GAS(no_top_gas_espana, Convert.ToInt32(fd.ToString("yyyyMM")),
                            "ES", "gas", "NO TOP", fh.Day);
                    }


                    #endregion

                    #region Cisternas gas España
                    Console.WriteLine("cisternas gas España");
                    workSheet = excelPackage.Workbook.Worksheets["cisternas gas España"];
                    List<EndesaEntity.facturacion.cuadroDeMando.InformeGas> cisternas_gas_espana =
                        dic_gas.Values.Where(z => z.pais == "España" &&
                        z.es_cisterna == true).ToList();


                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.InformeGas p in cisternas_gas_espana)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.id_PS; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.nombre_punto_suminsitro; c++;
                        workSheet.Cells[f, c].Value = p.cups; c++;
                        workSheet.Cells[f, c].Value = p.grupo; c++;
                        workSheet.Cells[f, c].Value = p.fecha_inicio;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        if (p.fecha_fin < new DateTime(4999, 12, 31))
                        {
                            workSheet.Cells[f, c].Value = p.fecha_fin;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.ultimo_mes_facturado; c++;

                        workSheet.Cells[f, c].Value = p.medida;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.fecha_emision_sigame;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;


                        workSheet.Cells[f, c].Value = p.fecha_informe;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        // DIF (F. INFORME - F. EMISIÓN)
                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_informe - p.fecha_emision_sigame).TotalDays);

                        }
                        c++;

                        if (p.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_emision_sigame - p.ffactura).TotalDays);
                            c++;

                        }
                        else
                        {
                            c++; // FFACTURA
                            c++; // DIF (FFACTURA - F.EMISIÓN)                        
                        }




                        workSheet.Cells[f, c].Value = p.pendiente; c++;
                        workSheet.Cells[f, c].Value = p.comentario; c++;
                        workSheet.Cells[f, c].Value = p.promedio_facturacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    if (mostrar_resumen)
                    {
                        resumen.Save_GAS(cisternas_gas_espana, Convert.ToInt32(fd.ToString("yyyyMM")),
                            "ES", "gas", "CISTERNAS", fh.Day);
                    }


                    #endregion

                    #region Gas Portugal
                    Console.WriteLine("Hoja gas Portugal");
                    workSheet = excelPackage.Workbook.Worksheets["gas portugal"];
                    List<EndesaEntity.facturacion.cuadroDeMando.InformeGas> gasPortugal =
                        dic_gas.Values.Where(z => z.pais == "Portugal").ToList();


                    f = 1;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.InformeGas p in gasPortugal)
                    {
                        c = 1;
                        f++;
                        workSheet.Cells[f, c].Value = p.id_PS; c++;
                        workSheet.Cells[f, c].Value = p.cif; c++;
                        workSheet.Cells[f, c].Value = p.nombre_punto_suminsitro; c++;
                        workSheet.Cells[f, c].Value = p.cups; c++;
                        workSheet.Cells[f, c].Value = p.fecha_inicio;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        if (p.fecha_fin < new DateTime(4999, 12, 31))
                        {
                            workSheet.Cells[f, c].Value = p.fecha_fin;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.ultimo_mes_facturado; c++;

                        workSheet.Cells[f, c].Value = p.medida;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.fecha_emision_sigame;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;


                        workSheet.Cells[f, c].Value = p.fecha_informe;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        // DIF (F. INFORME - F. EMISIÓN)
                        if (p.fecha_emision_sigame > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_informe - p.fecha_emision_sigame).TotalDays);

                        }
                        c++;

                        if (p.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                            workSheet.Cells[f, c].Value = Convert.ToInt32((p.fecha_emision_sigame - p.ffactura).TotalDays);
                            c++;

                        }
                        else
                        {
                            c++; // FFACTURA
                            c++; // DIF (FFACTURA - F.EMISIÓN)                        
                        }




                        workSheet.Cells[f, c].Value = p.pendiente; c++;
                        workSheet.Cells[f, c].Value = p.comentario; c++;
                        workSheet.Cells[f, c].Value = p.promedio_facturacion;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    if (mostrar_resumen)
                    {
                        resumen.Save_GAS(gasPortugal, Convert.ToInt32(fd.ToString("yyyyMM")),
                            "PT", "gas", "TOTAL", fh.Day);
                    }


                    #endregion
                }

                if (mostrar_resumen)
                {
                    Console.WriteLine("Hoja Resumen");
                    workSheet = excelPackage.Workbook.Worksheets["Resumen"];

                    // Cabecera
                    workSheet.Cells[3, 1].Value = utilFechas.ConvierteMes_a_Letra(fd)
                        + " " + fd.Year;


                    // Pintamos dias TOP (GAS)                   
                    f = 7;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }

                    // Pintamos dias NO TOP (GAS)
                    f = 13;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }

                    // Pintamos dias CISTERNAS (GAS)
                    f = 19;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }

                    // Pintamos dias TOTAL PORTUGAL (GAS)
                    f = 27;
                    c = 3;
                    for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(d.ToString("dd")); c++;
                    }


                    List<EndesaEntity.facturacion.cuadroDeMando.Resumen> lista_resumen;


                    #region GAS

                    lista_resumen = resumen.GetValor("ES", "gas", "TOP");

                    f = 7;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {

                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {

                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }

                    lista_resumen = resumen.GetValor("ES", "gas", "NO TOP");

                    f = 13;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {

                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {

                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }


                    lista_resumen = resumen.GetValor("ES", "gas", "CISTERNAS");

                    f = 19;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {

                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {

                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }


                    lista_resumen = resumen.GetValor("PT", "gas", "TOTAL");

                    f = 27;
                    foreach (EndesaEntity.facturacion.cuadroDeMando.Resumen p in lista_resumen)
                    {

                        c = 3;
                        f++;
                        for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                        {

                            if (p.dias[d.Day] >= 0)
                                workSheet.Cells[f, c].Value = p.dias[d.Day];
                            c++;
                        }

                    }

                    #endregion

                }

                excelPackage.SaveAs(nombreSalidaExcel);
                excelPackage = null;

                if (automatico)
                    EnvioCorreoInformeCuadroMando(nombreSalidaExcel.FullName);

                //ControlEstadoPuntoFEDA();
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("InformeExcelGas: " + ex.Message);
            }
        }

        private List<EndesaEntity.facturacion.cuadroDeMando.InformeGas> 
            Trata_Grupos_Especiales(Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.GrupoCliente> dic)
        {

            List<EndesaEntity.facturacion.cuadroDeMando.InformeGas> lista =
                new List<InformeGas>();

            try
            {
                foreach (KeyValuePair<string, EndesaEntity.facturacion.cuadroDeMando.GrupoCliente> p in dic)
                {
                    //EndesaEntity.facturacion.cuadroDeMando.InformeGas o =
                    //    dic_gas.Values.Where(z => z.cups == p.Key).Select()

                    foreach(KeyValuePair<int, EndesaEntity.facturacion.cuadroDeMando.InformeGas> pp in dic_gas)
                    {
                        if(pp.Value.cups == p.Key)
                        {
                            pp.Value.grupo = p.Value.grupo;
                            pp.Value.nombre_punto_suminsitro = p.Value.cliente;
                            lista.Add(pp.Value);
                        }
                    }


                    //if (o != null)
                    //{
                    //    o.grupo = p.Value.grupo;
                    //    o.nombre_punto_suminsitro = p.Value.cliente;
                    //    lista.Add(o);
                    //}
                    //else
                    //{
                    //    o = new EndesaEntity.facturacion.cuadroDeMando.InformeGas();
                    //    o.cups = p.Value.cups20;
                    //    o.grupo = p.Value.grupo;
                    //    o.nombre_punto_suminsitro = p.Value.cliente;
                    //}

                }

                return lista;
            }catch(Exception ex)
            {
                ficheroLog.AddError("Trata_Grupos_Especiales: " + ex.Message);
                if (lista.Count() > 0)
                    return lista;
                else
                    return new List<EndesaEntity.facturacion.cuadroDeMando.InformeGas>();
            }

            
        }

        private void EnvioCorreoInformeCuadroMando(string archivo)
        {
            
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                Console.WriteLine("Enviando correo Informe Cuadro de Mando Facturación");

                string from = param.GetValue("Buzon_envio_email");
                string to = param.GetValue("email_para");
                string cc = param.GetValue("email_copia");
                string subject = param.GetValue("email_asunto") + " " + DateTime.Now.ToString("dd/MM/yyyy");
                //string body = GeneraCuerpoHTML(CreaTabla(false), "No &Aacute;gora");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   Adjuntamos el informe de Cuadro de Mando de Facturación.");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);


                // EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (param.GetValue("EnviarMail") == "S")
                    mes.SendMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);

                else
                    mes.SaveMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("Buzon_envio_email"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreoInformeCuadroMando: " + e.Message);
            }
        }


        private void EnvioCorreoNoHayPendiente()
        {
                        
            StringBuilder textBody = new StringBuilder();

            try
            {
                Console.WriteLine("Enviando correo Informe Cuadro de Mando Facturación");

                string from = param.GetValue("Buzon_envio_email");
                //string to = param.GetValue("email_para");
                string to = "rsiope.gma@enel.com";
                //string cc = param.GetValue("email_copia");
                string subject = param.GetValue("email_asunto") + " " + DateTime.Now.ToString("dd/MM/yyyy");                

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   No se ha podido generar le cuadro de mando por indisponibilidad del Pendiente");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);


                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (param.GetValue("EnviarMail") == "S")
                    mes.SendMail(from, to, null, subject, textBody.ToString(), null);

                else
                    mes.SaveMail(from, to, null, subject, textBody.ToString(), null);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("Buzon_envio_email"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreoInformeCuadroMando: " + e.Message);
            }
        }

        // A petición de Angel (Junio 2024)
        // Función para controlar el cambio de estado del punto FEDA (ES9999000000000015AN)
        // El estado a monitorizar está parametrizado, pero inicialmente es: "01. Pendiente de medida"
        // Cuando cambie el estado a monitorizar a otro distinto a "FACTURADO" enviaremos durantes 3 días un email comunicando este cambio
        private void ControlEstadoPuntoFEDA()
        {
            //Variables locales, inicializamos a valores por defecto, que posteriormente se actualizan desde base de datos
            EndesaEntity.facturacion.cuadroDeMando.Informe o;
            string cups_feda = "ES9999000000000015AN";
            string estado_a_monitorizar = "01. Pendiente de medida";
            int monitor_estado_feda = 0;
            int dias_aviso_monitor_feda = 3;
            string estado_fin_aviso_monitor_feda = "FACTURADO";
            //string mail_to_monitor_estado_feda = "gbenavides@minsait.com";
            //string admin_mail_to_monitor_estado_feda = "gbenavides@minsait.com";
            //string enviar_mail_monitor_estado_feda = "N";
            string estado_feda_actual = "";
            
            
            try
            {
                //Obtener parámetros de base de datos
                cups_feda = param.GetValue("cups_feda");
                estado_a_monitorizar = param.GetValue("estado_a_monitorizar");
                monitor_estado_feda = Convert.ToInt32(param.GetValue("monitor_estado_feda"));
                dias_aviso_monitor_feda = Convert.ToInt32(param.GetValue("dias_aviso_monitor_feda"));
                //mail_to_monitor_estado_feda = param.GetValue("mail_to_monitor_estado_feda");
                //admin_mail_to_monitor_estado_feda = param.GetValue("admin_mail_to_monitor_estado_feda");
                //enviar_mail_monitor_estado_feda = param.GetValue("enviar_mail_monitor_feda");

                //Buscar CUPS20 FEDA en diccionario
                if (dic_electricidad_ES.TryGetValue(cups_feda, out o))
                {
                    estado_feda_actual = o.estado_LTP;
                    //Obtenemos estado_LTP y ejecutamos lógica control
                    if (estado_a_monitorizar == estado_feda_actual)
                    {
                        monitor_estado_feda = 0;
                    }
                    else
                    {
                        monitor_estado_feda++;
                    }
                    //Actualizamos parámetro en base de datos
                    param.UpdateParameter("monitor_estado_feda",monitor_estado_feda.ToString());


                    //Enviamos email si procede
                    if( (monitor_estado_feda > 0) && (monitor_estado_feda<=dias_aviso_monitor_feda) && (estado_feda_actual != estado_fin_aviso_monitor_feda))
                    {
                        //Enviamos email notificación facturación disponible para punto FEDA
                        EnvioCorreoEstadoPuntoFEDA(estado_feda_actual);
                    }
                    else
                    {
                        //Como control nos enviamos a un destinatario admin obtenido por parametro bbdd un email de estado actual punto FEDA
                        EnvioAdminCorreoEstadoPuntoFEDA(estado_feda_actual);
                    }
                }
                else
                {
                    //Atención: no se ha encontrado el punto FEDA, log y notificar a admin
                    ficheroLog.AddError("ControlEstadoPuntoFeda: no se ha encontrado el CUPS de FEDA");

                }
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("ControlEstadoPuntoFeda: " + ex.Message);
            }


        }
        private void EnvioCorreoEstadoPuntoFEDA(string estado_feda_actual)
        {

            StringBuilder textBody = new StringBuilder();
            
            try
            {
                Console.WriteLine("Enviando correo Notificación Estado Punto FEDA");

                string from = param.GetValue("Buzon_envio_email");
                //string to = param.GetValue("email_para");
                string to = param.GetValue("mail_to_monitor_estado_feda");
                string cc = param.GetValue("mail_cc_monitor_estado_feda");
                string subject = param.GetValue("asunto_mail_monitor_estado_feda") + " - " + DateTime.Now.ToString("dd/MM/yyyy");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días." : "Buenas tardes.");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Cliente: FEDA. La medida ya está disponible para facturarse: " + estado_feda_actual);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Saludos");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);


                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (param.GetValue("enviar_mail_monitor_estado_feda") == "S")
                    mes.SendMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), null, true);

                else
                    mes.SaveMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), null);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("Buzon_envio_email"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreoEstadoPuntoFEDA: " + e.Message);
            }
        }
        private void EnvioAdminCorreoEstadoPuntoFEDA(string estado_feda_actual)
        {

            StringBuilder textBody = new StringBuilder();

            try
            {
                Console.WriteLine("Enviando correo Notificación Estado Punto FEDA");

                string from = param.GetValue("Buzon_envio_email");
                //string to = param.GetValue("email_para");
                string to = param.GetValue("admin_mail_to_monitor_estado_feda");
                string cc = param.GetValue("admin_mail_to_monitor_estado_feda");
                string subject = "Estado punto FEDA - " + DateTime.Now.ToString("dd/MM/yyyy");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días." : "Buenas tardes.");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Cliente FEDA. Estado punto: " + estado_feda_actual);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Saludos");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);


                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                mes.SendMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), null, true);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("Buzon_envio_email"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioAdminCorreoEstadoPuntoFEDA: " + e.Message);
            }
        }
    }
}
