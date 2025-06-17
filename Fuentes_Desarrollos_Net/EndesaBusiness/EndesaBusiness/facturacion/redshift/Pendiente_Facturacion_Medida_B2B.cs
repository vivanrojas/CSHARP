using EndesaBusiness.utilidades;
using MySql.Data.MySqlClient;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.servidores;
using System.Drawing;
using System.Data.Odbc;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Aspose.Cells;
using Workbook = Aspose.Cells.Workbook;
using Worksheet = Aspose.Cells.Worksheet;

///using Worksheet = Aspose.Cells.Worksheet;
using System.Data;
using DataTable = System.Data.DataTable;
using Style = Aspose.Cells.Style;
using Range = Aspose.Cells.Range;
using EndesaEntity.global;
using System.Reflection;
using MySQLDB = EndesaBusiness.servidores.MySQLDB;

namespace EndesaBusiness.facturacion.redshift
{
    public class Pendiente_Facturacion_Medida_B2B
    {
        utilidades.Param param;
        utilidades.Seguimiento_Procesos ss_pp;
        logs.Log ficheroLog;

        Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> dic_pendiente_hist_fecha;
        Dictionary<string, DateTime> dic_dias_estado;

        EndesaBusiness.facturacion.redshift.Pendiente_Subestados subestados_sap;
        EndesaBusiness.medida.pendiente.PendienteWeb_B2B pendienteWeb_B2B;
        EndesaBusiness.medida.pendiente.PendienteWeb_B2B pendienteWeb_B2BBTN;
        EndesaBusiness.medida.EstadosSubestadosKronos subestadosKronos;
        EndesaBusiness.medida.EstadosSubestadosKronos subestadosKronosBTN;
        EndesaBusiness.medida.pendiente.PendienteFacturadoSAPKEE pendienteSAPKEE;
        EndesaBusiness.medida.pendiente.SegmentosPortugal_KEE SegmentosPortugal_KEE;
        EndesaBusiness.medida.pendiente.PendienteWeb_B2B RelacionIncidencias;
        EndesaEntity.medida.Relacion_Incidencias_SAP RelacionIncidenciasSAP;
        //EndesaBusiness.facturacion.redshift.Relacion_Incidencias_SAP RelacionIncidenciasSAP;
        EndesaBusiness.medida.pendiente.PendienteWeb_B2B FechaBajaSap;
        EndesaBusiness.medida.pendiente.PendienteWeb_B2B RelacionIncidenciasCUPS;

        string SubestadosNoEncontrados = "";

        public void Kill(bool entireProcessTree) { }

        public Pendiente_Facturacion_Medida_B2B()
        {
    
            param = new utilidades.Param("t_ed_h_pendiente_sap_kronos_param", servidores.MySQLDB.Esquemas.FAC);
            ss_pp = new utilidades.Seguimiento_Procesos();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_MED_Pendiente_Facturacion_Medida_B2B");

            string FechaPdteWebB2B_fhact;
            FechaPdteWebB2B_fhact = CalculaFechaMaxPdteWebB2B().ToString(); //fec_act

            if (Convert.ToDateTime(FechaPdteWebB2B_fhact).ToString("yyyy-MM-dd") != DateTime.Now.ToString("yyyy-MM-dd"))
            {
                ss_pp.Update_Comentario("Facturación", "Informe Pendiente BI", "Informe SAP - KRONOS", "No está actualizada la tabla ed_owner.t_ed_h_pdtweb_pm_b2b");
                //System.Windows.Forms.MessageBox.Show("No está actualizada la tabla ed_owner.t_ed_h_pdtweb_pm_b2b");
                ficheroLog.Add("No está actualizada la tabla ed_owner.t_ed_h_pdtweb_pm_b2");
                Kill(true);
            }
            else {

                subestados_sap = new Pendiente_Subestados();
                RelacionIncidenciasSAP = new EndesaEntity.medida.Relacion_Incidencias_SAP();
                //RelacionIncidenciasSAP = new facturacion.redshift.Relacion_Incidencias_SAP();
                subestadosKronos = new medida.EstadosSubestadosKronos();
                subestadosKronosBTN = new medida.EstadosSubestadosKronos("BTN");
                pendienteWeb_B2B = new medida.pendiente.PendienteWeb_B2B(); 
                pendienteWeb_B2BBTN = new medida.pendiente.PendienteWeb_B2B("BTN");
                FechaBajaSap= new medida.pendiente.PendienteWeb_B2B(1);
                //////RelacionIncidenciasCUPS= new medida.pendiente.PendienteWeb_B2B(2); // es de Mysql la carga (la carga en la biblioteca está hecha y funciona), no lo sigo desarrollando en el código, mysql no penaliza
            }

            ///RelacionIncidencias = new medida.pendiente.PendienteWeb_B2B(true);
            //////SegmentosPortugal_KEE = new medida.pendiente.SegmentosPortugal_KEE();

        }

        private DateTime UltimaActualizacionMySQL()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime fecha = new DateTime(2022, 01, 01);

            strSql = "SELECT max(fh_envio) AS fh_envio FROM t_ed_h_sap_pendiente_facturar";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
                if (r["fh_envio"] != System.DBNull.Value)
                    fecha = Convert.ToDateTime(r["fh_envio"]);
            db.CloseConnection();
            Console.WriteLine("Última fecha de copiado MySQL: " + fecha.ToString("dd/MM/yyyy"));
            ficheroLog.Add("Última fecha de copiado MySQL: " + fecha.ToString("dd/MM/yyyy"));

            return fecha;
        }

        public void GeneraInformePendSAP(bool automatico)
        {
            FileInfo file;
            string ruta_salida_archivo = "";

            string FechaPdteWebB2B_fhact;
            FechaPdteWebB2B_fhact = CalculaFechaMaxPdteWebB2B().ToString(); //fec_act

            if (Convert.ToDateTime(FechaPdteWebB2B_fhact).ToString("yyyy-MM-dd") != DateTime.Now.ToString("yyyy-MM-dd"))
            {
                System.Windows.Forms.MessageBox.Show("No está actualizada la tabla ed_owner.t_ed_h_pdtweb_pm_b2b");
                Kill(true);
            }
            else
            {
           
               string[] listaArchivos = System.IO.Directory.GetFiles(automatico ? param.GetValue("ruta_salida_informe") : @"c:\Temp\",
                        param.GetValue("prefijo_informe") + "*.xlsx");

                for (int i = 0; i < listaArchivos.Length; i++)
                {
                    file = new FileInfo(listaArchivos[i]);
                    file.Delete();
                }

                if (automatico)
                    ruta_salida_archivo = param.GetValue("ruta_salida_informe")
                        + param.GetValue("prefijo_informe")
                        + DateTime.Now.ToString("yyyyMMdd")
                        + param.GetValue("sufijo_informe");
                else
                    ruta_salida_archivo = @"c:\Temp\"
                       + param.GetValue("prefijo_informe")
                       + DateTime.Now.ToString("yyyyMMdd")
                       + param.GetValue("sufijo_informe");

                InformePendiente_BI_Facturacion(ruta_salida_archivo, automatico);

            }
        }

        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> CargaPendienteHist_DesdeFecha(DateTime f, List<string> lista_empresas)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>>();

            DateTime fecha_actual = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime fecha_registro = new DateTime();
            int meses_pdtes = 0;
            int aniomes = 0;
            int num_dias_informe = 0;
            bool firstOnly = true;
            string multipunto;
            bool blnRecalculoEstado;
            string areaincidencia;
            string subestado_incidencia;

            MySQLDB dbIncidencia;
            MySqlDataReader inci;
            bool ExisteIncidencia;
            string Auxiliar;
            int Conteo;
            string BuscaCadena = "";

            try
            {

                //sof.Contruye_Sofisticados();
                //agoraManual = CargaAgoraManual(DateTime.Now, DateTime.Now);
                //agoraPortugal = new contratacion.Agora_Portugal();

                /*strSql = " SELECT fecha_informe, pend.empresa_titular AS EMPRESA,"
                    + " pend.cups13, "
                    + " pend.mes as aaaammPdte, pend.estado, pend.subestado, pend.tam"
                    + " FROM fact. pend where "
                    + " fecha_informe >= '" + f.ToString("yyyy-MM-dd") + "'"
                    + " ORDER BY pend.fecha_informe, pend.empresa_titular, "
                    + " pend.cups13, pend.mes ASC";
                */

                strSql = "SELECT p.cd_empr_titular, ps.cd_empr, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                        //+ " ps.fh_alta_crto, ps.fh_inicio_vers_crto, ps.cups20, ps.cd_tarifa_c,"
                        + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, p.cd_cups as cups20, ps.cd_tarifa_c,"
                        + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre, p.cd_estado, p.cd_subestado,"
                        + " de.de_estado, ds.de_subestado, p.fh_periodo, IF(s.CUPS20 IS NULL, p.agora, 'S') AS agora, p.TAM,"
                        + " p.lg_multimedida, p.fec_act, p.fh_desde, p.fh_hasta, ps.fh_prev_fin_crto, ps.fh_baja_crto,p.fh_periodo as mes"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"

                        // 06/02/2025 añadir sofisticados agora
                        + " LEFT OUTER JOIN fact.cm_sofisticados s on s.CUPS20 = p.cd_cups"
                        //Fin 06/02/2025

                    ////+ " where p.fec_act >= '" + DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd") + "'"
                    //Paco 04/04/2024
                    //Hay problema de sincronización entre SAP y KEE, KEE se ejecuta todos los días y SAP no tiene pases los fines de semana/¿Festivos?
                    //los lunes puede esta más avanzado KEE que SAP, lanzamos con f menos 10 (f es las máxima fh_envio de sap_pendiente_facturar_agrupado)
                    //////+ " where p.fec_act >= '" + f.AddDays(-10).ToString("yyyy-MM-dd") + "'"
                    + " where p.fec_act >= '" + f.ToString("yyyy-MM-dd") + "'"
                //+ " and p.fec_act <= '" + f.ToString("yyyy-MM-dd") + "'"
                //+ " and p.cd_cups in ('ES0022000007936420TG')"
                ///

                //Hay que quitar (porque así se decidió en la Negociación de este reporte) que no pueden aparecer registros con el mes pendiente = al mes en curso !!!!!!!!!!!
                + " and p.fh_periodo <> '" + DateTime.Now.ToString("yyyyMM") + "'";
                //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!11


                foreach (string p in lista_empresas)
                {
                    if (firstOnly)
                    {
                        strSql += " and cd_empr_titular in ("
                            + "'" + p + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + p + "'";
                }

                strSql += ") ORDER BY p.fec_act desc, ps.cd_empr, "
                    //+ " ps.cups20, p.fh_periodo ASC";
                    +" p.cd_cups, p.fh_periodo ASC";


                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.GBL);

                //////command= new MySqlCommand("SET net_write_timeout=99999;", db.con);
                //////command= new MySqlCommand("SET net_read_timeout=99999;", db.con);
                //r = command.ExecuteNonQuery;
                MySqlCommand comm;
                using (comm = new MySqlCommand("set net_write_timeout=99999; set net_read_timeout=99999", db.con))
                {
                    comm.ExecuteNonQuery();
                }


                comm = new MySqlCommand(strSql, db.con);
                ////command.CommandTimeout = 60 * 60 * 2;
                r = comm.ExecuteReader();

                Conteo = 1;
                while (r.Read())
                {

                    EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                    c.cod_empresaTitular = r["cd_empr_titular"].ToString();
                    c.empresaTitular = r["cd_empr"].ToString();
                    c.cups20 = r["cups20"].ToString();

                    c.aaaammPdte = Convert.ToInt32(r["fh_periodo"]);
                    c.cod_estado = r["cd_estado"].ToString();
                    c.cod_subestado = r["cd_subestado"].ToString();
                    c.estado = r["de_estado"].ToString();
                    c.subsEstado = r["de_subestado"].ToString();
                    c.fecha_informe = Convert.ToDateTime(r["fec_act"]).Date;
                    if (r["fh_prev_fin_crto"] != System.DBNull.Value)
                    {
                        c.fh_prev_fin_crto = Convert.ToDateTime(r["fh_prev_fin_crto"]).Date;
                    }
                    if (r["fh_baja_crto"] != System.DBNull.Value)
                    {
                        c.fh_baja_crto = Convert.ToDateTime(r["fh_baja_crto"]).Date;
                    }

                    c.fh_desde = Convert.ToDateTime(r["fh_desde"]).Date;
                    c.fh_hasta = Convert.ToDateTime(r["fh_hasta"]).Date;

                    c.fh_act = Convert.ToDateTime(r["fec_act"]).Date;

                    if (r["tam"] != System.DBNull.Value)
                    {
                        if (c.aaaammPdte != 0)
                        {
                            aniomes = Convert.ToInt32(c.aaaammPdte);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);

                            //Paco 21/01/2025
                            //////meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12)
                            //////    + fecha_actual.Month - fecha_registro.Month;

                            meses_pdtes = ((c.fecha_informe.Year - fecha_registro.Year) * 12)
                              + c.fecha_informe.Month - fecha_registro.Month;
                            //  Fin Paco 21/01/2025

                            c.tam = Convert.ToDouble(r["tam"]) * meses_pdtes;
                        }
                        else
                            c.tam = Convert.ToDouble(r["tam"]);

                    }

                    if (r["agora"] != System.DBNull.Value)
                        if (r["agora"].ToString() != "N")
                            c.agora = true;
                        else
                            c.agora = false;
                    else
                        c.agora = false;

                    if (r["lg_multimedida"] != System.DBNull.Value)
                        c.multimedida = r["lg_multimedida"].ToString() == "S";
                    else
                        c.multimedida = false;

                    multipunto = "";

                    //Vemos si es o no multipunto
                    List<string> Multipuntos = new List<string>();
                    Multipuntos = pendienteWeb_B2B.GetCupsMultipunto(r["cups20"].ToString());
                    if (Multipuntos.Count > 0)
                    {
                        if (Multipuntos[0] == "False")
                        {
                            multipunto = "N";
                        }
                        else
                        {
                            multipunto = "S";
                        }

                    }
                    else
                    {
                        multipunto = "N";
                    }

                    // 11-04-2025
                    if (r["mes"] != System.DBNull.Value)
                    {  
                        aniomes = Convert.ToInt32(r["mes"]);
                    }

                    // Fin 11/04/2025
                   
                  
                    //if (r["cups20"].ToString() == "ES0021000001531208ZA") {

                    //    System.Windows.Forms.MessageBox.Show("");
                    //};

                    //AÑADIMOS las etiquetas de SAP
                    if (subestados_sap.AreaResponsableMedida(r["cd_subestado"].ToString()))
                    {
                        subestadosKronos.existe = false;
                        //he cambiado GetEstadoKEE  por GetEstadoKEEDetalle

                        //subestadosKronos.descripcion_subestado = "";
                        //subestadosKronos.temporal = "";
                        //c.cod_subestado = "";
                        //////c.cod_subestado_SAP = "";

                        //IMPORTANTE!!!! Limpiamos el temporal
                        subestadosKronos.temporal = "";
                        subestadosKronos.area_responsable=null;

                        if (multipunto == "S")
                        {
                            subestadosKronos.GetEstadoKEEDetalleMultipunto(pendienteWeb_B2B.GetCupsDetalle(r["cups20"].ToString(),
                                Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]), Convert.ToDateTime(r["fec_act"])));
                        }
                        else
                        {
                            subestadosKronos.GetEstadoKEEDetalle(pendienteWeb_B2B.GetCupsDetalle(r["cups20"].ToString(),
                                Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]), Convert.ToDateTime(r["fec_act"])));
                        }

                        //////subestadosKronos.GetEstadoKEEDetalle(pendienteWeb_B2B.GetCupsDetalle(r["cups20"].ToString(),
                        //////    Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]), Convert.ToDateTime(r["fec_act"])));

                        if (subestadosKronos.existe)
                        {
                            //ORIGINAL
                            //////c.subestado_SAP =  subestadosKronos.descripcion_subestado;

                            ////////workSheet.Cells[f, c].Value = subestadosKronos.descripcion_estado; c++;
                            //////c.cod_subestado_SAP = subestadosKronos.descripcion_subestado.Substring(0, subestadosKronos.descripcion_subestado.IndexOf(" "));
                            // FIN ORIGINAL

                            //11-04-2025
                            BuscaCadena = subestadosKronos.descripcion_estado;

                            if (BuscaCadena.IndexOf("Discrepancia") >= 0)
                            {
                                if (BuscaCadena.IndexOf("Periodo SAP encontrado parcialmente en el informe pendiente de KEE") >= 0)
                                {

                                    c.subestado_SAP = "01.B05 Error Sistemas KEE - SAP - Incoherencia periodos KEE-SAP";
                                    c.cod_subestado_SAP = "01.B05";

                                }

                                else
                                {
                                    if (BuscaCadena.IndexOf("Periodos del pendiente de KEE contenidos en las fechas de SAP") >= 0 || BuscaCadena.IndexOf("Periodo SAP contenido dentro de un informe pendiente de KEE") >= 0)
                                    {
                                        //////workSheet.Cells[f, c].Value = subestadosKronos.estado_periodo; c++;
                                        //////workSheet.Cells[f, c].Value = subestadosKronos.area_responsable; c++;
                                        //////// Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el subestado de KEE lo he guardado en temporal de momento
                                        //////workSheet.Cells[f, c].Value = subestadosKronos.temporal; c++;
                                        //////// Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el estado de KEE lo he guardado en descripcion_subestado
                                        ////////en lugar de en descripcion_estado
                                        //////workSheet.Cells[f, c].Value = subestadosKronos.descripcion_subestado; c++;
                                        //////string[] cadena = subestadosKronos.descripcion_estado.Split(';');
                                        //////string[] cadenaAux = cadena[2].Split('/');
                                        //////workSheet.Cells[f, c].Value = cadenaAux[0].Substring(1, 8).Substring(6, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(4, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(0, 4); c++;
                                        //////workSheet.Cells[f, c].Value = cadenaAux[1].Substring(0, 8).Substring(6, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(4, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(0, 4); c++;
                                        ///

                                        c.subestado_SAP = subestadosKronos.temporal;
                                        c.cod_subestado_SAP = subestadosKronos.temporal.Substring(0, subestadosKronos.temporal.IndexOf(" "));
                                    }
                                    else
                                    {

                                        if (BuscaCadena.IndexOf("No existe el cups en el informe del pendiente de KEE") >= 0)
                                        {

                                            c.subestado_SAP = "01.B07 Error Sistemas KEE - SAP - Pdte SAP - No existe en KEE";
                                            c.cod_subestado_SAP = "01.B07";

                                        }
                                        else
                                        {

                                            //////if (c.agora == true)
                                            //////{
                                            //////    MessageBox.Show(r["cups20"].ToString());
                                            //////}



                                            if (BuscaCadena.IndexOf("No existen fechas en el informe pendiente de KEE para el periodo") >= 0)
                                            {

                                                c.subestado_SAP = "01.B06 Error Sistemas KEE-SAP - Pdte SAP-No pdte K";
                                                c.cod_subestado_SAP = "01.B06";

                                            }
                                            else
                                            {
                                                c.subestado_SAP = "";
                                                c.cod_subestado_SAP = "";
                                            }

                                        } //Fin  if (BuscaCadena.IndexOf("No existe el cups en el informe del pendiente de KEE") >= 0)

                                    } // fin  if (BuscaCadena.IndexOf("Periodos del pendiente de KEE contenidos en las fechas de SAP") >= 0 || BuscaCadena.IndexOf("Periodo SAP contenido dentro de un informe pendiente de KEE") >= 0)

                                } // Fin  if (BuscaCadena.IndexOf("Periodo SAP encontrado parcialmente en el informe pendiente de KEE") >= 0 )


                                //////workSheet.Cells[f, c].Value = subestadosKronos.descripcion_estado; c++;

                            }
                            else
                            {
                                //////workSheet.Cells[f, c].Value = subestadosKronos.estado_periodo; c++;
                                //////workSheet.Cells[f, c].Value = subestadosKronos.area_responsable; c++;
                                //////workSheet.Cells[f, c].Value = subestadosKronos.descripcion_subestado; c++;
                                //////workSheet.Cells[f, c].Value = subestadosKronos.descripcion_estado; c++;
                               
                                c.subestado_SAP = subestadosKronos.descripcion_subestado;
                                c.cod_subestado_SAP = subestadosKronos.descripcion_subestado.Substring(0, subestadosKronos.descripcion_subestado.IndexOf(" "));
                            }
                            /// Fin 11-04-2025


                            //////if (subestadosKronos.temporal == "" || subestadosKronos.temporal == null)
                            //////{

                            //////    if (subestadosKronos.descripcion_estado.IndexOf("No existe el cups en el informe del pendiente de KEE") >= 0)
                            //////    {
                            //////        c.subestado_SAP = "01.B07 Error Sistemas KEE - SAP - Pdte SAP - No existe en KEE";
                            //////        c.cod_subestado_SAP = "01.B07";
                            //////    }
                            //////    else
                            //////    {
                            //////        if (subestadosKronos.descripcion_estado.IndexOf("No existen fechas en el informe pendiente de KEE para el periodo") >= 0)
                            //////        {
                            //////            c.subestado_SAP = "01.B06 Error Sistemas KEE-SAP - Pdte SAP-No pdte K";
                            //////            c.cod_subestado_SAP = "01.B06";
                            //////        }
                            //////        else
                            //////        {
                            //////            //11-04-2025
                            //////            if (subestadosKronos.descripcion_estado.IndexOf("Periodo SAP encontrado parcialmente en el informe pendiente de KEE") >= 0)
                            //////            {
                            //////                c.subestado_SAP = "01.B05 Error Sistemas KEE - SAP - Incoherencia periodos KEE-SAP";
                            //////                c.cod_subestado_SAP = "01.B05";
                            //////            }
                            //////            else

                            //////            {
                            //////                c.subestado_SAP = subestadosKronos.descripcion_subestado;
                            //////                c.cod_subestado_SAP = subestadosKronos.descripcion_subestado.Substring(0, subestadosKronos.descripcion_subestado.IndexOf(" "));
                            //////            }               
                            //////        }
                            //////    }

                            //////if (subestadosKronos.descripcion_estado == "Discrepancia: No existe el cups en el informe del pendiente de KEE")
                            //////{
                            //////    c.subestado_SAP = null;
                            //////    c.cod_subestado_SAP = null;
                            //////}
                            //////else
                            //////{
                            //////    c.subestado_SAP = subestadosKronos.descripcion_subestado;
                            //////    c.cod_subestado_SAP = subestadosKronos.descripcion_subestado.Substring(0, subestadosKronos.descripcion_subestado.IndexOf(" "));
                            //////}
                            //////}
                            //////else
                            //////{
                            //////    c.subestado_SAP = subestadosKronos.temporal;
                            //////    c.cod_subestado_SAP = subestadosKronos.temporal.Substring(0, subestadosKronos.temporal.IndexOf(" "));
                            //////}


                            ///Paco 14/05/2024 - Control de incidencias /////////////////////////////////////////////////
                            //Existe KEE, evaluo area de incidencia con area de KEE

                            if (aniomes > 0)
                            {
                                strSql = " select area from Relacion_INC_CUPS"
                                 + " where cups='" + r["cups20"].ToString() + "' and Mes_pendiente='" + aniomes.ToString() + "'"
                                 + " order by Fecha_apertura asc";

                                dbIncidencia = new MySQLDB(MySQLDB.Esquemas.MED);
                                command = new MySqlCommand(strSql, dbIncidencia.con);
                                inci = command.ExecuteReader();

                                ////ExisteIncidencia = false;

                                while (inci.Read())
                                {
                                    Auxiliar = Reincidente(inci["area"].ToString(), subestadosKronos.area_responsable);
                                    if (Auxiliar != "")
                                    {
                                        c.subestado_SAP = Auxiliar;
                                        c.cod_subestado_SAP = Auxiliar.Substring(0, 6);
                                    }
                                } // Fin  while (inci.Read())

                                dbIncidencia.CloseConnection();
                            }
                            ////////////////////////////////////////////////////////////////////////////////////////////////
                        }
                    }

                    // Sólo necesitamos los últimos 5 días para el informe                    
                    List<EndesaEntity.medida.Pendiente> o;
                    if (!d.TryGetValue(c.fecha_informe, out o))
                    {

                        o = new List<EndesaEntity.medida.Pendiente>();
                        o.Add(c);
                        d.Add(c.fecha_informe, o);
                    }
                    else
                        o.Add(c);


                    ///System.Diagnostics.Debug.WriteLine(Conteo.ToString() + "-" + c.cups20 + "-" + c.fh_act + "-" + c.cod_subestado + "-" + c.cod_subestado + "-" + c.cod_subestado_SAP + "-" + c.cod_subestado_SAP + Conteo.ToString());
                    System.Diagnostics.Debug.WriteLine(Conteo.ToString());
                    Conteo = Conteo + 1;
                    //////GC.Collect();
                    //////GC.WaitForPendingFinalizers();

                }
                db.CloseConnection();
                r.Close();
                return d;
            }
            catch (Exception e)
            {
                ///MessageBox.Show(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                ficheroLog.addError("CargaPendiente: " + e.Message + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                return null;
            }
        }
        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> CargaPendienteHist_PT_DesdeFecha(DateTime f, List<string> lista_empresas, List<string> lista_segmentos)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>>();

            DateTime fecha_actual = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime fecha_registro = new DateTime();
            int meses_pdtes = 0;
            int aniomes = 0;
            int num_dias_informe = 0;
            bool firstOnly = true;


            MySQLDB dbIncidencia;
            MySqlDataReader inci;
            bool ExisteIncidencia;
            string Auxiliar;
            string Segmentos;
            string BuscaCadena = "";

            try
            {
                //sof.Contruye_Sofisticados();
                //agoraManual = CargaAgoraManual(DateTime.Now, DateTime.Now);
                //agoraPortugal = new contratacion.Agora_Portugal();


                strSql = "SELECT p.cd_empr_titular, ps.cd_empr, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                        + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, p.cd_cups, ps.cd_tarifa_c, "
                        + " case when p.cd_segmento_ptg is null then segmento ELSE p.cd_segmento_ptg END AS cd_segmento_ptg, "
                        + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre, p.cd_estado, p.cd_subestado,"
                        + " de.de_estado, ds.de_subestado, p.fh_periodo, IF(s.CUPS20 IS NULL, p.agora, 'S') AS agora, p.TAM,"
                        + " p.lg_multimedida, p.fec_act, p.fh_desde, p.fh_hasta, p.fh_periodo as mes"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps_pt ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"


                        // 06/02/2025 añadir sofisticados agora
                        + " LEFT OUTER JOIN fact.cm_sofisticados s on s.CUPS20 = p.cd_cups"
                  //Fin 06/02/2025

                  ////+ " where p.fec_act >= '" + DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd") + "'";
                  //+ " and ps.cups20 in ('PT0002000111610262EM')"; 
                  ///07-11-2024 Para recuperar los segmentos de mercado nulo
                  + " LEFT JOIN fact.t_ed_h_sap_pendiente_facturar_segmentos_compor"
                  + " ON p.cd_cups = t_ed_h_sap_pendiente_facturar_segmentos_compor.cd_cups "
                /// Fin 07/11/2024 Para recuperar los segmentos de mercado nulo
                //Hay problema de sincronización entre SAP y KEE, KEE se ejecuta todos los días y SAP no tiene pases los fines de semana/¿Festivos?
                //los lunes puede esta más avanzado KEE que SAP, lanzamos con f menos 10 (f es las máxima fh_envio de sap_pendiente_facturar_agrupado)
                ///+ " where p.fec_act >= '" + f.AddDays(-10).ToString("yyyy-MM-dd") + "'"
                + " where p.fec_act >= '" + f.ToString("yyyy-MM-dd") + "'"

                //////+ " and p.cd_cups in ('PT0002000111610262EM','PT0002000116086035HR')"
                ///


                // 27/11/2024
                + " AND (p.fh_desde >= t_ed_h_sap_pendiente_facturar_segmentos_compor.fdesde "
                + " AND t_ed_h_sap_pendiente_facturar_segmentos_compor.fdesde <= p.fh_hasta) "
                // Fin 27/11/2024

                //Hay que quitar (porque así se decidió en la Negociación de este reporte) que no pueden aparecer registros con el mes pendiente = al mes en curso !!!!!!!!!!!
                + " and p.fh_periodo <> '" + DateTime.Now.ToString("yyyyMM") + "'";
                //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!11

                foreach (string p in lista_empresas)
                {
                    if (firstOnly)
                    {
                        strSql += " and cd_empr_titular in ("
                            + "'" + p + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + p + "'";
                }

                firstOnly = true;
                Segmentos = "";
                foreach (string p in lista_segmentos)
                {
                    if (firstOnly)
                    {
                        strSql += ") and (cd_segmento_ptg in ("
                            + "'" + p + "'";
                        firstOnly = false;
                        Segmentos= "'" + p + "'";
                    }
                    else
                    {
                        strSql += ",'" + p + "'";
                        Segmentos = Segmentos + ",'" + p + "'";
                    }
                        
                }

                //11-04-2025
                //////strSql += ") or cd_segmento_ptg is null ) ORDER BY p.fec_act desc, ps.cd_empr, "
                //////     + " ps.cups20, p.fh_periodo ASC";

                strSql += ") or cd_segmento_ptg is null )";

                strSql += " Group by  p.cd_empr_titular, ps.cd_empr, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                     + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, p.cd_cups, ps.cd_tarifa_c, "
                     + " case when p.cd_segmento_ptg is null then segmento ELSE p.cd_segmento_ptg END  , "
                     + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre, p.cd_estado, p.cd_subestado,"
                     + " de.de_estado, ds.de_subestado, p.fh_periodo, IF(s.CUPS20 IS NULL, p.agora, 'S') , p.TAM,"
                     + " p.lg_multimedida, p.fec_act, p.fh_desde, p.fh_hasta, p.fh_periodo"
                     + " ORDER BY p.fec_act desc, ps.cd_empr, ps.cups20, p.fh_periodo ASC";

                // Fin 11-04-2025

                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.GBL);

                MySqlCommand comm;
                using (comm = new MySqlCommand("set net_write_timeout=99999; set net_read_timeout=99999", db.con))
                {
                    comm.ExecuteNonQuery();
                }

                comm = new MySqlCommand(strSql, db.con);
                ////command.CommandTimeout = 60 * 60 * 2;
                r = comm.ExecuteReader();

                //////command = new MySqlCommand(strSql, db.con);
                //////r = command.ExecuteReader();
                while (r.Read())
                {

                    if (r["cd_segmento_ptg"] != System.DBNull.Value)
                    {
                        if (Segmentos.IndexOf(r["cd_segmento_ptg"].ToString()) > 0)
                        {

                            EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                            c.cod_empresaTitular = r["cd_empr_titular"].ToString();
                            c.empresaTitular = r["cd_empr"].ToString();
                            c.cups20 = r["cd_cups"].ToString();
                            c.aaaammPdte = Convert.ToInt32(r["fh_periodo"]);
                            c.cod_estado = r["cd_estado"].ToString();
                            c.cod_subestado = r["cd_subestado"].ToString();
                            c.estado = r["de_estado"].ToString();
                            c.subsEstado = r["de_subestado"].ToString();
                            c.fecha_informe = Convert.ToDateTime(r["fec_act"]).Date;

                            c.fh_desde = Convert.ToDateTime(r["fh_desde"]).Date;
                            c.fh_hasta = Convert.ToDateTime(r["fh_hasta"]).Date;

                            if (r["cd_segmento_ptg"] != System.DBNull.Value)
                                c.segmento = r["cd_segmento_ptg"].ToString();


                            if (r["tam"] != System.DBNull.Value)
                            {
                                if (c.aaaammPdte != 0)
                                {
                                    aniomes = Convert.ToInt32(c.aaaammPdte);
                                    fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                        Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                                    // Paco 21/01/2025
                                    ////meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12)
                                    ////    + fecha_actual.Month - fecha_registro.Month;
                                    ////c.tam = Convert.ToDouble(r["tam"]) * meses_pdtes;

                                    meses_pdtes = ((c.fecha_informe.Year - fecha_registro.Year) * 12)
                                    + c.fecha_informe.Month - fecha_registro.Month;

                                    c.tam = Convert.ToDouble(r["tam"]) * meses_pdtes;
                                    // Paco 21/01/2025
                                }
                                else
                                    c.tam = Convert.ToDouble(r["tam"]);
                            }

                            if (r["agora"] != System.DBNull.Value)
                                if (r["agora"].ToString() != "N")
                                    c.agora = true;
                                else
                                    c.agora = false;
                            else
                                c.agora = false;

                            //if (r["lg_multimedida"] != System.DBNull.Value)
                            //    c.multimedida = r["lg_multimedida"].ToString() == "S";
                            //else
                            //    c.multimedida = false;

                            // 11-04-2025
                            if (r["mes"] != System.DBNull.Value)
                            {
                                aniomes = Convert.ToInt32(r["mes"]);
                            }

                            //AÑADIMOS las etiquetas de SAP
                            if (subestados_sap.AreaResponsableMedida(r["cd_subestado"].ToString()))
                            {
                                //OLD
                                //////subestadosKronos.GetEstadoKEE(pendienteWeb_B2B.GetCupsDetalle(r["cups20"].ToString(),
                                //////    Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]), Convert.ToDateTime(r["fec_act"])));
                                //FIN OLD

                                //IMPORTANTE!!!! Limpiamos el temporal
                                subestadosKronos.temporal = "";
                                subestadosKronos.area_responsable = null;


                                subestadosKronos.GetEstadoKEEDetalle(pendienteWeb_B2B.GetCupsDetalle(r["cd_cups"].ToString(),
                                    Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]), Convert.ToDateTime(r["fec_act"])));




                                //OLD
                                //////if (subestadosKronos.existe)
                                //////{
                                //////    c.subestado_SAP = subestadosKronos.descripcion_subestado;
                                //////    c.cod_subestado_SAP = subestadosKronos.descripcion_subestado.Substring(0, subestadosKronos.descripcion_subestado.IndexOf(" "));
                                //////}
                                //OLD
                                if (subestadosKronos.existe)
                                {

                                    ////////////if (subestadosKronos.temporal == "" || subestadosKronos.temporal == null)
                                    ////////////{

                                    ////////////    if (subestadosKronos.descripcion_estado == "Discrepancia: No existe el cups en el informe del pendiente de KEE")
                                    ////////////    {
                                    ////////////        c.subestado_SAP = "01.B07 Error Sistemas KEE - SAP - Pdte SAP - No existe en KEE";
                                    ////////////        c.cod_subestado_SAP = "01.B07";
                                    ////////////    }
                                    ////////////    else
                                    ////////////    {
                                    ////////////        if (subestadosKronos.descripcion_estado.IndexOf("No existen fechas en el informe pendiente de KEE para el periodo") >= 0)
                                    ////////////        {
                                    ////////////            c.subestado_SAP = "01.B06 Error Sistemas KEE-SAP - Pdte SAP-No pdte K";
                                    ////////////            c.cod_subestado_SAP = "01.B06";
                                    ////////////        }
                                    ////////////        else
                                    ////////////        {
                                    ////////////            //11-04-2025
                                    ////////////            if (subestadosKronos.descripcion_estado.IndexOf("Periodo SAP encontrado parcialmente en el informe pendiente de KEE") >= 0)
                                    ////////////            {
                                    ////////////                c.subestado_SAP = "01.B05 Error Sistemas KEE - SAP - Incoherencia periodos KEE-SAP";
                                    ////////////                c.cod_subestado_SAP = "01.B05";
                                    ////////////            }
                                    ////////////            else

                                    ////////////            {
                                    ////////////                c.subestado_SAP = subestadosKronos.descripcion_subestado;
                                    ////////////                c.cod_subestado_SAP = subestadosKronos.descripcion_subestado.Substring(0, subestadosKronos.descripcion_subestado.IndexOf(" "));
                                    ////////////            }
                                    ////////////            //////c.subestado_SAP = subestadosKronos.descripcion_subestado;
                                    ////////////            //////c.cod_subestado_SAP = subestadosKronos.descripcion_subestado.Substring(0, subestadosKronos.descripcion_subestado.IndexOf(" "));
                                    ////////////        }
                                    ////////////    }
                                    ////////////}
                                    ////////////else
                                    ////////////{
                                    ////////////    c.subestado_SAP = subestadosKronos.temporal;
                                    ////////////    c.cod_subestado_SAP = subestadosKronos.temporal.Substring(0, subestadosKronos.temporal.IndexOf(" "));
                                    ////////////}

                                    //11-04-2025
                                    BuscaCadena = subestadosKronos.descripcion_estado;

                                    if (BuscaCadena.IndexOf("Discrepancia") >= 0)
                                    {
                                        if (BuscaCadena.IndexOf("Periodo SAP encontrado parcialmente en el informe pendiente de KEE") >= 0)
                                        {
                                            c.subestado_SAP = "01.B05 Error Sistemas KEE - SAP - Incoherencia periodos KEE-SAP";
                                            c.cod_subestado_SAP = "01.B05";
                                        }
                                        else
                                        {
                                            if (BuscaCadena.IndexOf("Periodos del pendiente de KEE contenidos en las fechas de SAP") >= 0 || BuscaCadena.IndexOf("Periodo SAP contenido dentro de un informe pendiente de KEE") >= 0)
                                            {
                                                //////workSheet.Cells[f, c].Value = subestadosKronos.estado_periodo; c++;
                                                //////workSheet.Cells[f, c].Value = subestadosKronos.area_responsable; c++;
                                                //////// Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el subestado de KEE lo he guardado en temporal de momento
                                                //////workSheet.Cells[f, c].Value = subestadosKronos.temporal; c++;
                                                //////// Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el estado de KEE lo he guardado en descripcion_subestado
                                                ////////en lugar de en descripcion_estado
                                                //////workSheet.Cells[f, c].Value = subestadosKronos.descripcion_subestado; c++;
                                                //////string[] cadena = subestadosKronos.descripcion_estado.Split(';');
                                                //////string[] cadenaAux = cadena[2].Split('/');

                                                c.subestado_SAP = subestadosKronos.temporal;
                                                c.cod_subestado_SAP = subestadosKronos.temporal.Substring(0, subestadosKronos.temporal.IndexOf(" "));
                                            }
                                            else
                                            {

                                                if (BuscaCadena.IndexOf("No existe el cups en el informe del pendiente de KEE") >= 0)
                                                {

                                                    c.subestado_SAP = "01.B07 Error Sistemas KEE - SAP - Pdte SAP - No existe en KEE";
                                                    c.cod_subestado_SAP = "01.B07";

                                                }
                                                else
                                                {

                                                    //////if (c.agora == true)
                                                    //////{
                                                    //////    MessageBox.Show(r["cups20"].ToString());
                                                    //////}

                                                    if (BuscaCadena.IndexOf("No existen fechas en el informe pendiente de KEE para el periodo") >= 0)
                                                    {

                                                        c.subestado_SAP = "01.B06 Error Sistemas KEE-SAP - Pdte SAP-No pdte K";
                                                        c.cod_subestado_SAP = "01.B06";

                                                    }
                                                    else
                                                    {
                                                        c.subestado_SAP = "";
                                                        c.cod_subestado_SAP = "";
                                                    }

                                                } //Fin  if (BuscaCadena.IndexOf("No existe el cups en el informe del pendiente de KEE") >= 0)

                                            } // fin  if (BuscaCadena.IndexOf("Periodos del pendiente de KEE contenidos en las fechas de SAP") >= 0 || BuscaCadena.IndexOf("Periodo SAP contenido dentro de un informe pendiente de KEE") >= 0)

                                        } // Fin  if (BuscaCadena.IndexOf("Periodo SAP encontrado parcialmente en el informe pendiente de KEE") >= 0 )


                                        //////workSheet.Cells[f, c].Value = subestadosKronos.descripcion_estado; c++;

                                    }
                                    else
                                    {
                                        //////workSheet.Cells[f, c].Value = subestadosKronos.estado_periodo; c++;
                                        //////workSheet.Cells[f, c].Value = subestadosKronos.area_responsable; c++;
                                        //////workSheet.Cells[f, c].Value = subestadosKronos.descripcion_subestado; c++;
                                        //////workSheet.Cells[f, c].Value = subestadosKronos.descripcion_estado; c++;

                                        c.subestado_SAP = subestadosKronos.descripcion_subestado;
                                        c.cod_subestado_SAP = subestadosKronos.descripcion_subestado.Substring(0, subestadosKronos.descripcion_subestado.IndexOf(" "));
                                    }
                                    /// Fin 11-04-2025


                                    ///Paco 14/05/2024 - Control de incidencias /////////////////////////////////////////////////
                                    //Existe KEE, evaluo area de incidencia con area de KEE
                                    if (aniomes > 0)
                                    {
                                        strSql = " select area from Relacion_INC_CUPS"
                                         + " where cups='" + r["cd_cups"].ToString() + "' and Mes_pendiente='" + aniomes.ToString() + "'"
                                         + " order by Fecha_apertura asc";

                                        dbIncidencia = new MySQLDB(MySQLDB.Esquemas.MED);
                                        command = new MySqlCommand(strSql, dbIncidencia.con);
                                        inci = command.ExecuteReader();

                                        ////ExisteIncidencia = false;

                                        while (inci.Read())
                                        {
                                            Auxiliar = Reincidente(inci["area"].ToString(), subestadosKronos.area_responsable);
                                            if (Auxiliar != "")
                                            {
                                                c.subestado_SAP = Auxiliar;
                                                c.cod_subestado_SAP = Auxiliar.Substring(0, 6);
                                            }
                                        } // Fin  while (inci.Read())

                                        dbIncidencia.CloseConnection();
                                    }

                                }

                            }

                            // Sólo necesitamos los últimos 5 días para el informe                    
                            List<EndesaEntity.medida.Pendiente> o;
                            if (!d.TryGetValue(c.fecha_informe, out o))
                            {

                                o = new List<EndesaEntity.medida.Pendiente>();
                                o.Add(c);
                                d.Add(c.fecha_informe, o);
                            }
                            else
                                o.Add(c);
                        }// Fin
                    } // Fin if (r["cd_segmento_ptg"] != System.DBNull.Value)

                } // Fin while (r.Read())

                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                ficheroLog.addError("CargaPendiente: " + e.Message);
                return null;
            }
        }




        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> CargaPendienteHist_PT_DesdeFecha_BTN(DateTime f, List<string> lista_empresas, List<string> lista_segmentos)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>>();

            DateTime fecha_actual = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime fecha_registro = new DateTime();
            int meses_pdtes = 0;
            int aniomes = 0;
            int num_dias_informe = 0;
            bool firstOnly = true;
            int NumeroDias;
            DateTime fecha_calculada = new DateTime();
            string Segmentos;


            MySQLDB dbIncidencia;
            MySqlDataReader inci;
            bool ExisteIncidencia;
            string Auxiliar;
            string BuscaCadena = "";
            string cups = "";

            try
            {
                //sof.Contruye_Sofisticados();
                //agoraManual = CargaAgoraManual(DateTime.Now, DateTime.Now);
                //agoraPortugal = new contratacion.Agora_Portugal();


                strSql = "SELECT p.cd_empr_titular, ps.cd_empr, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                        + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, p.cd_cups, ps.cd_tarifa_c, "
                        ////+ " p.cd_segmento_ptg,"
                        // 07/11/2024  Para recuperar los segmentos de mercado nulo
                        + " case when p.cd_segmento_ptg is null then segmento ELSE p.cd_segmento_ptg END AS cd_segmento_ptg, "
                        // // 07/11/2024 Para recuperar los segmentos de mercado nulo
                        + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre, p.cd_estado, p.cd_subestado,"
                        + " de.de_estado, ds.de_subestado, p.fh_periodo, IF(s.CUPS20 IS NULL, p.agora, 'S') AS agora, p.TAM,"
                        + " p.lg_multimedida, p.fec_act, p.fh_desde, p.fh_hasta,p.fh_periodo as mes"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps_pt ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"
                        ////+ " where p.fec_act >= '" + DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd") + "'";

                        // 06/02/2025 añadir sofisticados agora
                        + " LEFT OUTER JOIN fact.cm_sofisticados s on s.CUPS20 = p.cd_cups"
                  //Fin 06/02/2025

                  //Hay problema de sincronización entre SAP y KEE, KEE se ejecuta todos los días y SAP no tiene pases los fines de semana/¿Festivos?
                  //los lunes puede esta más avanzado KEE que SAP, lanzamos con f menos 10 (f es las máxima fh_envio de sap_pendiente_facturar_agrupado)
                  ///+ " where p.fec_act >= '" + f.AddDays(-10).ToString("yyyy-MM-dd") + "'"
                  ///
                  ///07-11-2024 Para recuperar los segmentos de mercado nulo
                  + " LEFT JOIN fact.t_ed_h_sap_pendiente_facturar_segmentos_compor"
                  + " ON p.cd_cups = t_ed_h_sap_pendiente_facturar_segmentos_compor.cd_cups "
                /// Fin 07/11/2024 Para recuperar los segmentos de mercado nulo
                /// 
                /// + " where p.fec_act >= '" + f.AddDays(-10).ToString("yyyy-MM-dd") + "'"
                ///07-11-2024 REcupero unicamente los últimos cinco días de fec_activacion, que son los que pinto.
                + " where p.fec_act >= '" + f.ToString("yyyy-MM-dd") + "'"

                //////+ " and p.cd_cups in ( 'PT0002000023467508BJ','PT0002000004930716FW')"


                /// Fin 07-11-2024 REcupero unicamente los últimos cinco días de fec_activacion, que son los que pinto.
                // 27/11/2024
                + " AND (p.fh_desde >= t_ed_h_sap_pendiente_facturar_segmentos_compor.fdesde "
                + " AND t_ed_h_sap_pendiente_facturar_segmentos_compor.fdesde <= p.fh_hasta) "
                // Fin 27/11/2024

                //Hay que quitar (porque así se decidió en la Negociación de este reporte) que no pueden aparecer registros con el mes pendiente = al mes en curso !!!!!!!!!!!
                + " and p.fh_periodo <> '" + DateTime.Now.ToString("yyyyMM") + "'";
                //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!11

                //////+ " and p.cd_cups in ('PT0002000028173825AR')";


                foreach (string p in lista_empresas)
                {
                    if (firstOnly)
                    {
                        strSql += " and cd_empr_titular in ("
                            + "'" + p + "'";
                        firstOnly = false;

                    }
                    else
                    {
                        strSql += ",'" + p + "'";

                    }

                }

                firstOnly = true;
                //////foreach (string p in lista_segmentos)
                //////{
                //////    if (firstOnly)
                //////    {
                //////        strSql += ") and cd_segmento_ptg in ("
                //////            + "'" + p + "'";
                //////        firstOnly = false;
                //////    }
                //////    else
                //////        strSql += ",'" + p + "'";
                //////}



                //////strSql += ") ORDER BY p.fec_act desc, ps.cd_empr, "
                //////     + " p.cd_cups, p.fh_periodo ASC";

                // 07/11/2024  controlos segmentos de mercado nulo, que ahora los recupero
                Segmentos = "";
                foreach (string p in lista_segmentos)
                {
                    if (firstOnly)
                    {
                        strSql += ") and (cd_segmento_ptg in ("
                            + "'" + p + "'";
                        firstOnly = false;
                        Segmentos = "'" + p + "'";
                    }
                    else
                    {
                        strSql += ",'" + p + "'";
                        Segmentos = Segmentos + ",'" + p + "'";
                    }

                }

                strSql += ") or cd_segmento_ptg  is null ) ";

                // Fin 07/11/2024

                // 11-04-2025
                strSql += " group by p.cd_empr_titular, ps.cd_empr, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                       + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, p.cd_cups, ps.cd_tarifa_c, "
                       + " case when p.cd_segmento_ptg is null then segmento ELSE p.cd_segmento_ptg END , "
                       + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre, p.cd_estado, p.cd_subestado,"
                       + " de.de_estado, ds.de_subestado, p.fh_periodo, IF(s.CUPS20 IS NULL, p.agora, 'S') , p.TAM,"
                       + " p.lg_multimedida, p.fec_act, p.fh_desde, p.fh_hasta,p.fh_periodo"
                       + "  ORDER BY p.fec_act desc, ps.cd_empr, p.cd_cups, p.fh_periodo ASC ";


                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.GBL);

                MySqlCommand comm;
                using (comm = new MySqlCommand("set net_write_timeout=99999; set net_read_timeout=99999", db.con))
                {
                    comm.ExecuteNonQuery();
                }

                comm = new MySqlCommand(strSql, db.con);
                ////command.CommandTimeout = 60 * 60 * 2;
                r = comm.ExecuteReader();

                //////command = new MySqlCommand(strSql, db.con);
                //////r = command.ExecuteReader();
                ///
                while (r.Read())
                {
                    //07/11/2024 de los nulos puedo haber recuperado alguno que no pertenece al segmento de mercado de la hoja a rellenar
                    if (r["cd_segmento_ptg"] != System.DBNull.Value)
                    {
                        if (Segmentos.IndexOf(r["cd_segmento_ptg"].ToString()) > 0)
                        {
                            // Fin 7/11/2024 de los nulos puedo haber recuperado alguno que no pertenece al segmento de mercado de la hoja a rellenar

                            EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                            c.cod_empresaTitular = r["cd_empr_titular"].ToString();
                            c.empresaTitular = r["cd_empr"].ToString();
                            c.cups20 = r["cd_cups"].ToString();

                            cups= r["cd_cups"].ToString();

                            c.aaaammPdte = Convert.ToInt32(r["fh_periodo"]);
                            c.cod_estado = r["cd_estado"].ToString();
                            c.cod_subestado = r["cd_subestado"].ToString();
                            c.estado = r["de_estado"].ToString();
                            c.subsEstado = r["de_subestado"].ToString();
                            c.fecha_informe = Convert.ToDateTime(r["fec_act"]).Date;

                            c.fh_desde = Convert.ToDateTime(r["fh_desde"]).Date;
                            c.fh_hasta = Convert.ToDateTime(r["fh_hasta"]).Date;

                            if (r["cd_segmento_ptg"] != System.DBNull.Value)
                                c.segmento = r["cd_segmento_ptg"].ToString();


                            if (r["tam"] != System.DBNull.Value)
                            {
                                if (c.aaaammPdte != 0)
                                {
                                    aniomes = Convert.ToInt32(c.aaaammPdte);
                                    fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                        Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                                    // Paco 21/01/2025
                                    //////meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12)
                                    //////    + fecha_actual.Month - fecha_registro.Month;
                                    //////c.tam = Convert.ToDouble(r["tam"]) * meses_pdtes;

                                    meses_pdtes = ((c.fecha_informe.Year - fecha_registro.Year) * 12)
                                   + c.fecha_informe.Month - fecha_registro.Month;

                                    c.tam = Convert.ToDouble(r["tam"]) * meses_pdtes;
                                    // Fin 21/01/2025

                                }
                                else
                                    c.tam = Convert.ToDouble(r["tam"]);

                            }

                            if (r["agora"] != System.DBNull.Value)
                                if (r["agora"].ToString() != "N")
                                    c.agora = true;
                                else
                                    c.agora = false;
                            else
                                c.agora = false;

                            //if (r["lg_multimedida"] != System.DBNull.Value)
                            //    c.multimedida = r["lg_multimedida"].ToString() == "S";
                            //else
                            //    c.multimedida = false;


                            // 11-04-2025
                            if (r["mes"] != System.DBNull.Value)
                            {
                                aniomes = Convert.ToInt32(r["mes"]);
                            }

                            //AÑADIMOS las etiquetas de SAP
                            if (subestados_sap.AreaResponsableMedida(r["cd_subestado"].ToString()))
                            {
                                //OLD
                                //////subestadosKronos.GetEstadoKEE(pendienteWeb_B2B.GetCupsDetalle(r["cups20"].ToString(),
                                //////    Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]), Convert.ToDateTime(r["fec_act"])));
                                //FIN OLD

                                /// 
                                //IMPORTANTE!!!! Limpiamos el temporal
                                subestadosKronosBTN.temporal = "";
                                subestadosKronosBTN.descripcion_subestado = "";
                                subestadosKronosBTN.descripcion_estado = "";
                                subestadosKronosBTN.area_responsable = null;
                                subestadosKronosBTN.cod_estado = "";

                                subestadosKronosBTN.GetEstadoKEEDetalleBTN(pendienteWeb_B2BBTN.GetCupsDetalle_BTN(r["cd_cups"].ToString(),
                                    Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]), Convert.ToDateTime(r["fec_act"])));

                                if (subestadosKronosBTN.existe)
                                {

                                    //////if (subestadosKronosBTN.temporal == "" || subestadosKronosBTN.temporal == null)
                                    //////{
                                    BuscaCadena = subestadosKronosBTN.descripcion_estado;


                                    if (BuscaCadena.IndexOf("Discrepancia") >= 0)
                                    {
                                        if (BuscaCadena.IndexOf("No existe el cups en el informe del pendiente de KEE") >= 0)
                                        {
                                            //////NumeroDias = 0;

                                            //////string[] cadena = subestadosKronosBTN.descripcion_estado.Split(';');
                                            //////string[] cadenaAux = cadena[1].Split('/');
                                            ////////cadenaAux[0].Substring(1, 8).Substring(6, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(4, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(0, 4); c++;

                                            ////////////802 Facturada DISTRIBUIDORA       Pendiente KEE   01.C25  Pendiente Medida KEE - En ciclo de Medida           01_DISTRIBUIDORA      01_DISTRIBUIDORA
                                            ////////////802 Facturada MEDIDA              Pendiente KEE   01.C25  Pendiente Sistemas KEE - Pte ejecución algoritmo    02_INCIDENCIA_MEDIDA  02_INCIDENCIA

                                            //////fecha_calculada = Convert.ToDateTime(cadenaAux[0].Substring(1, 8).Substring(6, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(4, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(0, 4));
                                            //////fecha_calculada = fecha_calculada.AddDays(-1);
                                            //////NumeroDias = (DateTime.Now - fecha_calculada).Days;


                                            NumeroDias = 0;
                                            NumeroDias = (DateTime.Now - Convert.ToDateTime(r["fh_desde"])).Days;

                                            //Controlo si la fecha desde -1 hasta el día de hoy supera los 90 días
                                            if (NumeroDias > 45)
                                            {
                                                c.subestado_SAP = "01.C25 Pendiente Sistemas KEE - Pte ejecución algoritmo";
                                                c.cod_subestado_SAP = "01.C25";
                                            }
                                            else
                                            {
                                                c.subestado_SAP = "01.C25 Pendiente Medida KEE - En ciclo de Medida";
                                                c.cod_subestado_SAP = "01.C25";
                                            }
                                        }
                                        else
                                        {
                                            if (BuscaCadena.IndexOf("No existen fechas en el informe pendiente de KEE para el periodo") >= 0)
                                            {
                                                NumeroDias = 0;

                                                string[] cadena = subestadosKronosBTN.descripcion_estado.Split(';');
                                                string[] cadenaAux = cadena[1].Split('/');
                                                //cadenaAux[0].Substring(1, 8).Substring(6, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(4, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(0, 4); c++;

                                                //////802 Facturada DISTRIBUIDORA       Pendiente KEE   01.C25  Pendiente Medida KEE - En ciclo de Medida           01_DISTRIBUIDORA      01_DISTRIBUIDORA
                                                //////802 Facturada MEDIDA              Pendiente KEE   01.C25  Pendiente Sistemas KEE - Pte ejecución algoritmo    02_INCIDENCIA_MEDIDA  02_INCIDENCIA

                                                fecha_calculada = Convert.ToDateTime(cadenaAux[0].Substring(1, 8).Substring(6, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(4, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(0, 4));
                                                fecha_calculada = fecha_calculada.AddDays(-1);
                                                NumeroDias = (DateTime.Now - fecha_calculada).Days;

                                                //Controlo si la fecha desde -1 hasta el día de hoy supera los 90 días
                                                if (NumeroDias > 45)
                                                {
                                                    c.subestado_SAP = "01.C25 Pendiente Sistemas KEE - Pte ejecución algoritmo";
                                                    c.cod_subestado_SAP = "01.C25";
                                                }
                                                else
                                                {
                                                    c.subestado_SAP = "01.C25 Pendiente Medida KEE - En ciclo de Medida";
                                                    c.cod_subestado_SAP = "01.C25";
                                                }
                                            }
                                            else
                                            {

                                                c.subestado_SAP = "";
                                                c.cod_subestado_SAP = "";
                                            }
                                        }
                                    }

                                    else // else de if (BuscaCadena.IndexOf("Discrepancia") >= 0)
                                    {
                                        if (subestadosKronosBTN.cod_estado == "802")
                                        {
                                            if ((DateTime.Now - subestadosKronosBTN.fh_hasta).Days > 45)
                                            {
                                                c.subestado_SAP = "01.C25 Pendiente Sistemas KEE - Pte ejecución algoritmo";
                                            }
                                            else
                                            {
                                                c.subestado_SAP = "01.C25 Pendiente Medida KEE - En ciclo de Medida";
                                            }
                                        }
                                        else
                                        {
                                            c.subestado_SAP = subestadosKronosBTN.descripcion_subestado;
                                        }
                                        c.cod_subestado_SAP = subestadosKronosBTN.descripcion_subestado.Substring(0, subestadosKronosBTN.descripcion_subestado.IndexOf(" "));

                                    } // fin  if (BuscaCadena.IndexOf("Discrepancia") >= 0)

                                    ///Paco 14/05/2024 - Control de incidencias /////////////////////////////////////////////////
                                    //Existe KEE, evaluo area de incidencia con area de KEE
                                    if (aniomes > 0)
                                    {
                                        strSql = " select area from Relacion_INC_CUPS"
                                         + " where cups='" + r["cd_cups"].ToString() + "' and Mes_pendiente='" + aniomes.ToString() + "'"
                                         + " order by Fecha_apertura asc";

                                        dbIncidencia = new MySQLDB(MySQLDB.Esquemas.MED);
                                        command = new MySqlCommand(strSql, dbIncidencia.con);
                                        inci = command.ExecuteReader();

                                        ////ExisteIncidencia = false;

                                        while (inci.Read())
                                        {
                                            if (inci["area"].ToString()!= "" & subestadosKronosBTN.existe == true)
                                            {             
                                                Auxiliar = Reincidente(inci["area"].ToString(), subestadosKronosBTN.area_responsable);
                                                if (Auxiliar != "")
                                                {
                                                    Console.WriteLine(cups);
                                                    c.subestado_SAP = Auxiliar;
                                                    c.cod_subestado_SAP = Auxiliar.Substring(0, 6);
                                                }
                                            }

                                        } // Fin  while (inci.Read())

                                        dbIncidencia.CloseConnection();
                                    }
                                } // Fin  if (subestadosKronosBTN.existe)

                            } // Fin   if (subestados_sap.AreaResponsableMedida(r["cd_subestado"].ToString()))

                            // Sólo necesitamos los últimos 5 días para el informe                    
                            List<EndesaEntity.medida.Pendiente> o;

                            if (!d.TryGetValue(c.fecha_informe, out o))
                            {

                                o = new List<EndesaEntity.medida.Pendiente>();
                                o.Add(c);
                                d.Add(c.fecha_informe, o);
                            }
                            else
                                o.Add(c);

                        } // Fin  if (Segmentos.IndexOf(r["cd_segmento_ptg"].ToString()) >0)
                    } //  if (r["cd_segmento_ptg"] != System.DBNull.Value) 
                }  // Fin  while (r.Read())
                db.CloseConnection();
                return d;
            } //Fin Try
            catch (Exception e)
            {
                ficheroLog.addError("CargaPendiente: " + e.Message);
                return null;
            }
        }


        private void InformePendiente_BI_Facturacion(string ruta_salida_archivo, bool automatico)
        {

            MySQLDB db;
            servidores.RedShiftServer dbRS;
            MySQLDB dbIncidencia;
            MySQLDB dbReporte;
            MySqlCommand command;
            MySqlDataReader r;
            MySqlDataReader rMinima;
            MySqlDataReader inci;
            MySqlDataReader reporte;
            string strSql = "";
            string strSqlEsqFact = "";
            string strSqlHoja = "";

            //Paco, para controlar el detalle, cuando no hay datos en t_ed_h_ps y hay que buscar en el histórico
            MySQLDB dbAux;
            MySqlCommand commandAux;
            MySqlDataReader rAux;
            string empresa;;
            string nif;
            string cliente;
            string apellido;
            DateTime? FechaAlta;
            DateTime? FechaInicio;
            string Tarifa;
            string contrato;
            string Distribuidora;
            string CadenaAuxiliar;
            //////////////////////////////////////////////////
            string ruta_salida_archivo_BTE;
            string ruta_salida_archivo_BTN;
     
            OdbcCommand commandRS;
            OdbcDataReader rRS;

            int c = 1;
            int f = 1;
            SubestadosNoEncontrados = "";

            DateTime fecha_actual = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime fecha_registro = new DateTime();
            int meses_pdtes = 0;
            int aniomes = 0;
          

            bool tiene_complemento_a01 = false;
            bool sacar_portugal = true;

            DateTime fd = new DateTime();
            DateTime fd_tam = new DateTime();
            DateTime udh = new DateTime();

            //MIO
            bool Pinto = true;
            string RangoInterno;
            string RangoPintoGris;
            int tamaño;
            string tamañoanterior = "";
            string cups20;
            string BuscaCadena ;

            //Paco 08/02/2024
            int PosicionBajaSap;
            int PosicionBajaKEE;
            int Cuenta;
            string agora;
            int PosicionSubEstadoSAP;
            int PosicionSubEstadoKEE;
            int PosicionSubEstadoGlobal;

            //Paco 13/03/2024 (para comprobación con tabla Relacion_INC_CUPS
            string areaincidencia;
            string incidencia;
            string estado_incidencia;
            string fecha_apertura;
            string prioridad_negocio;
            string titulo;
            string e_s_estado;
            string subestado_incidencia;
            string reincidente;
            Boolean ExisteIncidencia;
            Boolean EstadosReporte;
            string multipunto;
            DateTime fecha_calculada = new DateTime();

            utilidades.Fechas utilfecha = new Fechas();

            string Auxiliar;
            int DiasExistenciaKEE = 0;

            Boolean blnLanzarDiasKEE;
            MySQLDB dbDiasKEE;
            MySqlDataReader DiasKEE;
            MySQLDB dbDiasKEEMax;
            MySqlDataReader DiasKEEMax;

            int DiasExistenciaKEEGlobal = 0;
            MySQLDB dbDiasKEEGlobal;
            MySqlDataReader DiasKEEGobal;
            MySQLDB dbDiasKEEMaxGlobal;
            MySqlDataReader DiasKEEMaxGlobal;

            string FechaSapPendienteFacturar_fhenvio;

            DateTime FechaSapPendienteFacturar_Desdefhenvio;
         
            DateTime FechaMinimaPintarResumen;
            string FechaPdteWebB2B_fhact;
            string FechaPdteWebB2B_fhInforme;
            int intColumna=0;
            string IncidenciaFacturacion = "";
            string Estado_FAC_SE = "";
            string Titulo_FAC = "";

            string PintarEstado = "";
            string PintarSubestado = "";
            ExcelRange Etiqueta;
            int TotalPruebaES;
            int TotalPruebaMTBTE;
            int TotalPruebaBTN;
            int Dimension;
            int CuadroInicio;
            int CuadroFin;
            String BordeCuadro;
            String EtiquetaEstado = "";
            int TotalParcial;
            int TotalCuadro;
            int TotalGeneral;
            string Piloto = "";
            Boolean ExistenDatosCuadro;

            String Segmentos = "";
            String NombreHoja = "";

            string Agora = "";
            string Estado = "";
            string Subestado = "";
            DateTime DiaPintar;
            int FilaCambioAgora=0;          
            int FilaCambioResponsable;
            int TotalConteoResponsable=0;
            int TotalResponsableEconomico=0;
            string DiaAux = "";
            
            Boolean CambioAgora = false;
            Boolean CambioEstado = false;
            Boolean CambioSubEstado = false;
            Boolean PintoDias = true;
            int t;

            //Variables para el cuadro Resumen
            int TotalAgoraDia1 = 0;
            int TotalAgoraDia2 = 0;
            int TotalAgoraDia3 = 0;
            int TotalAgoraDia4 = 0;
            int TotalAgoraDia5 = 0;
            double TotalImporteDia1 = 0;
            double TotalImporteDia2 = 0;
            double TotalImporteDia3 = 0;
            double TotalImporteDia4 = 0;
            double TotalImporteDia5 = 0;
            int TotalAgoraDia1Final = 0;
            int TotalAgoraDia2Final = 0;
            int TotalAgoraDia3Final = 0;
            int TotalAgoraDia4Final = 0;
            int TotalAgoraDia5Final = 0;
            double TotalImporteDia1Final = 0;
            double TotalImporteDia2Final = 0;
            double TotalImporteDia3Final = 0;
            double TotalImporteDia4Final = 0;
            double TotalImporteDia5Final = 0;
            DateTime Dia1Date;
            DateTime Dia2Date;
            DateTime Dia3Date;
            DateTime Dia4Date;
            DateTime Dia5Date;
            String Dia1 = "";
            String Dia2= "";
            String Dia3 = "";
            String Dia4 = "";
            String Dia5= "";
            int b = 0;
            

            try
            {


                if (!automatico || (UltimaActualizacionMySQL().Date >
                    ss_pp.GetFecha_FinProceso("Facturación", "Informe Pendiente BI", "Informe Pendiente BI").Date))
                {

                    if (automatico)
                        ss_pp.Update_Fecha_Inicio("Facturación", "Informe Pendiente BI", "Informe Pendiente BI");


                    /*
                    FileInfo plantillaExcel =  new FileInfo(System.Environment.CurrentDirectory +  param.GetValue("plantilla_informe_pendiente"));

                    FileInfo fileInfo = new FileInfo(ruta_salida_archivo);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                    ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);

                    var workSheet = excelPackage.Workbook.Worksheets["Resumen ES"];
                    */

                    // *****************************************

                    // Crea un fichero excel dinamicamente
                    //ruta_salida_archivo = "c:\\temp\\DetallePendFact_SAP_20231128_Inventarios_ES_PT_TAM.xlsx";
                    FileInfo fileInfo = new FileInfo(ruta_salida_archivo);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                    ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                    //////////////Creo la primera hoja
                    //////var workSheet = excelPackage.Workbook.Worksheets.Add("Resumen ES");
                    //////var headerCells = workSheet.Cells[1, 1, 1, 17];
                    //////var headerFont = headerCells.Style.Font;

                    List<string> lista_empresas_ES = new List<string>();
                    lista_empresas_ES.Add("ES21");
                    lista_empresas_ES.Add("ES22");

                    List<string> lista_empresas_PT = new List<string>();
                    lista_empresas_PT.Add("PT1Q");

                    List<string> lista_segmentos_MT_BTE = new List<string>();
                    lista_segmentos_MT_BTE.Add("MT");
                    lista_segmentos_MT_BTE.Add("BTE");
                    // Por petición de Ignacio Villar se añade AT y MAT 25/02/2025                    
                    lista_segmentos_MT_BTE.Add("AT");
                    lista_segmentos_MT_BTE.Add("MAT");

                    List<string> lista_segmentos_BTN = new List<string>();
                    lista_segmentos_BTN.Add("BTN");

                    // Tomamos lo últimos 5 días hábiles
                    // Si se lanza el listado en día fin de semana
                    // hay que quitar un día más porque el viernes todavía no se ha procesado

                    FechaSapPendienteFacturar_fhenvio = "";
                    FechaPdteWebB2B_fhact = "";
                    FechaPdteWebB2B_fhInforme = "";
                    FechaSapPendienteFacturar_fhenvio = CalculaFechaDesdeinformeMaxPendiente().ToString(); //fh_envio

                    FechaSapPendienteFacturar_Desdefhenvio= CalculaFechaDesdeinformeMaxPendienteDesdeFechaEnvio(); 
                    FechaPdteWebB2B_fhact =CalculaFechaMaxPdteWebB2B().ToString(); //fec_act
                    FechaPdteWebB2B_fhInforme = CalculaFechaMaxPdteWebB2B_fhinforme(CalculaFechaDesdeinformeMaxPendiente()).ToString();

                    if (!utilfecha.EsLaborable())
                    {
                        fd = utilfecha.UltimoDiaHabilAnterior(
                               utilfecha.UltimoDiaHabilAnterior(
                                   utilfecha.UltimoDiaHabilAnterior(
                                       utilfecha.UltimoDiaHabilAnterior(
                                           utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil())))));

                        udh = utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil());
                    }
                    else
                    {
                        fd = utilfecha.UltimoDiaHabilAnterior(
                                utilfecha.UltimoDiaHabilAnterior(
                                    utilfecha.UltimoDiaHabilAnterior(
                                        utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil()))));

                        udh = utilfecha.UltimoDiaHabil();
                    }

                    /*
                    //dic_pendiente_hist_fecha = CargaPendienteHist_DesdeFecha(CalculaFechaDesdeinforme(),lista_empresas_ES);

                    fd_tam = utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil());
                    */
                    int totales_dia = 0;
                    double totales_dia_tam = 0;

                    bool noAgora = false;
                    bool siAgora = true;

                    // Nuevos Totales

                    int totales_noagora_01 = 0;
                    double totales_noagora_tam_01 = 0;

                    int totales_noagora_02 = 0;
                    double totales_noagora_tam_02 = 0;

                    int totales_noagora_03 = 0;
                    double totales_noagora_tam_03 = 0;

                    int totales_noagora_04 = 0;
                    double totales_noagora_tam_04 = 0;

                    int totales_noagora_05 = 0;
                    double totales_noagora_tam_05 = 0;

                    int totales_agora_01 = 0;
                    double totales_agora_tam_01 = 0;

                    int totales_agora_02 = 0;
                    double totales_agora_tam_02 = 0;

                    int totales_agora_03 = 0;
                    double totales_agora_tam_03 = 0;

                    int totales_agora_04 = 0;
                    double totales_agora_tam_04 = 0;

                    int totales_agora_05 = 0;
                    double totales_agora_tam_05 = 0;

                    

                    Dictionary<DateTime, int> dic_Totales_cups = new Dictionary<DateTime, int>();
                    Dictionary<DateTime, double> dic_Totales_tam = new Dictionary<DateTime, double>();

                    dic_dias_estado = CargaDiasEstado();

                    int dia = 0;
                    int dia_tam = 0;
                    int fila;
                    int columna;
                    int PosicionPegarSubEstadoGlobal;
                    int NumeroDias;
                    int PosicionReincidente;
                    int PosicionDiasEstadoKEE;
                    int NumeroDiasKEE = 0;
                    Boolean ExisteIncidenciaBTN = false;

                    int intConteoES = 0;
                    int intConteoMTBTE = 0;
                    int intConteoBTN = 0;
                    int intConteoESDetalle = 0;
                    int intConteoMTBTEDetalle = 0;
                    int intConteoBTNDetalle = 0;

                    ///dic_pendiente_hist_fecha = CargaPendienteHist_DesdeFecha(CalculaFechaDesdeinforme(), lista_empresas_ES);

                    fd_tam = utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil());

                    //********************************************************************************************************************
                    //MIO*****************************************************************************************************************
                    //********************************************************************************************************************

                    int InicioRango;
                    int pintoetiqueta;
                    bool blnPintoEtiqueta;
                    bool blnRecalculoEstado;

                    blnPintoEtiqueta = true;
                    int Hoja;
                    List<string> listaEmpresas = new List<string>();
                    List<string> listaSegmentos = new List<string>();
                    List<string> listaAltaKEE = new List<string>();
                    List<string> fhBajaSap = new List<string>();
                    List<string> listaIncidenciasSAP = new List<string>();

                   
                    Color colorDeCeldaTitulo = ColorTranslator.FromHtml("#b3e3af");
                    Color colorDeCeldaCabecera = ColorTranslator.FromHtml("#c1dce6");
                    Color colorDeCeldaTotales = ColorTranslator.FromHtml("#e0f5c6");
                    Color colorDeCeldaGris = ColorTranslator.FromHtml("#e8eced");
                    ///ExcelRange Etiqueta;

                    //////pendienteWeb_B2B = new medida.pendiente.PendienteWeb_B2B();
                    ///
                    //////SegmentosPortugal_KEE = new medida.pendiente.SegmentosPortugal_KEE();


                    //********************************************************************************************************************
                    //FIN MIO ************************************************************************************************************
                    //********************************************************************************************************************
                    DateTime fecha_informe = CalculaFechaDetalle();
                    //////subestadosKronos = new medida.EstadosSubestadosKronos();
                    //////subestados_sap = new Pendiente_Subestados();
                    //////pendienteWeb_B2B = new medida.pendiente.PendienteWeb_B2B();
                    pendienteSAPKEE = new medida.pendiente.PendienteFacturadoSAPKEE();

                    blnLanzarDiasKEE = false;
                    strSql = "select max(FH_ACTUALIZACION) as FHMAX from kee_Dias";
                    dbDiasKEEMax = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql, dbDiasKEEMax.con);
                    DiasKEEMax = command.ExecuteReader();
                    while (DiasKEEMax.Read())
                    {
                        if (Convert.ToDateTime(DiasKEEMax["FHMAX"]).ToString("yyyy-MM-dd")== DateTime.Now.ToString("yyyy-MM-dd"))
                        {
                            blnLanzarDiasKEE = false;
                        }
                        else
                        {

                            blnLanzarDiasKEE = true;
                        }
                    }
                    dbDiasKEEMax.CloseConnection();
                    //Convert.ToDateTime(r["fh_desde"]).Date;
                    f = 1;

                    #region Detalle ES


                    ////////Creo la primera hoja
                    var workSheet = excelPackage.Workbook.Worksheets.Add("Detalle ES");
                    var headerCells = workSheet.Cells[1, 1, 1, 43];
                    var headerFont = headerCells.Style.Font;

                    //////workSheet = excelPackage.Workbook.Worksheets.Add("Detalle ES");
                    //////headerCells = workSheet.Cells[1, 1, 1, 43];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;
                    var allCells = workSheet.Cells[1, 1, 50, 50];

                    f = 1;
                    c = 1;
                    
                    workSheet.View.FreezePanes(2, 1);

                    headerFont.Bold = true;
                    workSheet.Cells[f, c].Value = "CUPS20"; c++;
                    workSheet.Cells[f, c].Value = "PERIODO"; c++;
                    workSheet.Cells[f, c].Value = "MES"; c++;
                    workSheet.Cells[f, c].Value = "FH_DESDE"; c++;
                    workSheet.Cells[f, c].Value = "FH_HASTA"; c++;
                    PosicionPegarSubEstadoGlobal = c;
                    workSheet.Cells[f, c].Value = "IMPORTE PDTE FACTURAR"; c++;
                    workSheet.Cells[f, c].Value = "MESES PDTES FACTURAR"; c++;
                    workSheet.Cells[f, c].Value = "ÁGORA"; c++;
                    workSheet.Cells[f, c].Value = "CLIENTE"; c++;
                    workSheet.Cells[f, c].Value = "DIAS_ESTADO_SAP"; c++;
                    PosicionDiasEstadoKEE = c;
                    workSheet.Cells[f, c].Value = "DIAS_ESTADO_KEE"; c++;
                    workSheet.Cells[f, c].Value = "DIAS_ESTADO_GLOBAL"; c++;
                    workSheet.Cells[f, c].Value = "Incidencia Facturacion"; c++;
                    workSheet.Cells[f, c].Value = "Estado_Fac_SE"; c++;
                    workSheet.Cells[f, c].Value = "Titulo_Fac"; c++;
                    workSheet.Cells[f, c].Value = "Incidencia Medida"; c++;
                    PosicionReincidente = c;
                    workSheet.Cells[f, c].Value = "Reincidente"; c++;
                    workSheet.Cells[f, c].Value = "Estado incidencia"; c++;
                    workSheet.Cells[f, c].Value = "Fecha_Apertura"; c++;
                    workSheet.Cells[f, c].Value = "Prioridad"; c++;
                    workSheet.Cells[f, c].Value = "Titulo"; c++;
                    workSheet.Cells[f, c].Value = "E_S_Estado"; c++;
                    workSheet.Cells[f, c].Value = "FH_ALTA_SALESFORCE"; c++;
                    workSheet.Cells[f, c].Value = "FH_ALTA_KEE"; c++;
                    workSheet.Cells[f, c].Value = "FH_ALTA_SAP"; c++; //FH_ALTA_SAP
                    workSheet.Cells[f, c].Value = "FH_BAJA_SALESFORCE"; c++;
                    PosicionBajaSap = c;
                    workSheet.Cells[f, c].Value = "FH_BAJA_SAP"; c++;
                    workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                    workSheet.Cells[f, c].Value = "TIPO CLIENTE"; c++;
                    workSheet.Cells[f, c].Value = "NIF"; c++;
                    workSheet.Cells[f, c].Value = "FPSERCON"; c++;
                    workSheet.Cells[f, c].Value = "Nº INSTALACIÓN"; c++;
                    workSheet.Cells[f, c].Value = "TARIFA"; c++;
                    workSheet.Cells[f, c].Value = "CONTRATO"; c++;  
                    workSheet.Cells[f, c].Value = "DISTRIBUIDORA"; c++;
                    workSheet.Cells[f, c].Value = "ESTADO"; c++;
                    PosicionSubEstadoSAP = c;
                    workSheet.Cells[f, c].Value = "SUBESTADO"; c++;
                    ///workSheet.Cells[f, c].Value = "DIAS_ESTADO"; c++;
                    workSheet.Cells[f, c].Value = "TAM"; c++;
                    
                    workSheet.Cells[f, c].Value = "ULT FH DESDE FACTURADA"; c++;
                    workSheet.Cells[f, c].Value = "ULT FH HASTA FACTURADA"; c++; 
                    workSheet.Cells[f, c].Value = "Estado periodo KEE"; c++;
                    workSheet.Cells[f, c].Value = "Área responsable KEE"; c++;
                    PosicionSubEstadoKEE = c;
                    workSheet.Cells[f, c].Value = "Subestado KEE"; c++;
                    workSheet.Cells[f, c].Value = "Estado KEE"; c++;
                    workSheet.Cells[f, c].Value = "FH_DESDE_KEE"; c++;
                    workSheet.Cells[f, c].Value = "FH_HASTA_KEE"; c++;
                    workSheet.Cells[f, c].Value = "Fecha BAJA KEE"; c++;
                    PosicionBajaKEE = c;
                    workSheet.Cells[f, c].Value = "Discrepancias"; c++;
                    PosicionSubEstadoGlobal = c;
                    workSheet.Cells[f, c].Value = "Subestado global"; c++;  
                    workSheet.Cells[f, c].Value = "ESTADO GLOBAL"; c++;
                    workSheet.Cells[f, c].Value = "ESTADO GLOBAL A REPORTAR"; c++;
                    workSheet.Cells[f, c].Value = "AREA ESTADO"; c++;
                    workSheet.Cells[f, c].Value = "multipunto"; c++;

                    //////workSheet.Cells[f, c].Value = "Nº DÍAS ESTADO KEE"; c++;

                    //for (int i = 1; i <= c; i++)
                    //{
                    //    workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    //}


                    int FhAltaSap;//Para quedarme con la posición de FH_ALTA_SAP
                    FhAltaSap = 0;
                    int intCliente;//Para quedarme con la posición de Cliente (nombre y apellidos)
                    intCliente =0;
                 
                    strSql = DetalleExcel(fecha_informe, lista_empresas_ES, null);

                    Boolean Existetedhps; //Para buscar los que vienen en blanco

                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        f++;
                        Debug.WriteLine(f);
                        c = 1;
                        empresa = ""; ;
                        nif = "";
                        cliente = "";
                        apellido = "";
                        FechaAlta = null;
                        FechaInicio = null;
                        Tarifa = "";
                        contrato = "";
                        Distribuidora = "";
                        cups20 = "";

                        multipunto = "";
                        NumeroDiasKEE = 0;
                    
                        List<string> Multipuntos = new List<string>();
                        Multipuntos = pendienteWeb_B2B.GetCupsMultipunto(r["cups20"].ToString());
                        if (Multipuntos.Count > 0)
                        {
                            if (Multipuntos[0] == "False")
                            {
                                multipunto = "N";
                            }
                            else
                            {
                                multipunto = "S";
                            }

                        }
                        else {
                             multipunto = "N";
                        }

                        if (r["cups20"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["cups20"].ToString();
                            cups20 = r["cups20"].ToString();
                        }
                        c++;

                        // Periodo
                        workSheet.Cells[f, c].Value = ""; c++;

                        if (r["mes"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["mes"]);
                            aniomes = Convert.ToInt32(r["mes"]);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                        }
                        c++;

                        if (r["fh_desde"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_desde"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["fh_hasta"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_hasta"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        // importe pdte facturar
                        if (r["mes"] != System.DBNull.Value && r["TAM"] != System.DBNull.Value)
                        {
                            //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["tam"]) * meses_pdtes;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            c++;
                        }
                        else
                        {
                            c++;
                        }

                        // meses pdte facturar
                        if (r["mes"] != System.DBNull.Value)
                        {
                            //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                            workSheet.Cells[f, c].Value = meses_pdtes;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        c++;

                        if (meses_pdtes > 1)
                        {
                            workSheet.Cells[f, 2].Value = ">1P";
                        }
                        else
                        {
                            workSheet.Cells[f, 2].Value = "1P";
                        }

                        if (r["agora"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["agora"].ToString();

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        //Cliente -- me quedo con la posición, para pintarlo más abajo
                        workSheet.Cells[f, c].Value = "";
                        intCliente = c;
                        c++ ;

                        ///DIAS_ESTADO_SAP
                        if (r["cups20"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = GetDiasEstado(r["cups20"].ToString());
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        c++;

                        //DIAS_ESTADO_KEE-- lo dejo sin rellenar, se pinta al final
                        workSheet.Cells[f, c].Value = "";
                        c++;

                        //DIAS_ESTADO_GLOBAL-- lo dejo sin rellenar, se pinta al final
                        workSheet.Cells[f, c].Value = "";
                        c++;


                        //Recuperar Incidencia Facturacion
                        listaIncidenciasSAP = RelacionIncidenciasSAP.GetCups(r["cups20"].ToString() + aniomes.ToString());

                        if (listaIncidenciasSAP.Count == 0)
                        {
                            IncidenciaFacturacion = "";
                            Estado_FAC_SE = "";
                            Titulo_FAC = "";
                        }
                        else
                        {
                            foreach (var mm in listaIncidenciasSAP)
                            {
                                string[] cadena = mm.Split(';');
                                IncidenciaFacturacion = cadena[0];
                                Estado_FAC_SE = cadena[1];
                                Titulo_FAC = cadena[2];
                                break;
                            }
                        }
                        //Fin recuperar incidencia facturacion
                        
                        // Comprobamos con la tabla Relacion_INC_CUPS
                        ////// 1.Se relaciona el detalle del informe obtenido(campos cups20 y mes) con la tabla "Relacion_INC_CUPS"(campos cups y mes).Si para un cups se encuentra en esa relación una incidencia que afeca al mismo mes que está pendiente se deberá desechar el estado que tuviera el informe y actualizarlo por uno nuevo.
                        ////// 2.El estado a actualizar dependerá del campo "area" de la tabla "Relacion_INC_CUPS", pudiendo existir 3 opciones:
                        ////// --01.B09 Incidencia_Contratacion: se informará con este estado cuando encuentre para ese mes una incidencia en el area de contratacion, independientemente de que también haya otras para ese mismo mes en otras areas.
                        //////-- 01.B10 Incidencia_Medida: no estando en el caso anterior, se informará con este estado cuando encuentre para ese mes una incidencia en el area de medida, independientemente de que también haya otras para ese mismo mes en el area de facturacion.
                        ////// --01.B11 Incidencia_Facturacion: no estando en ninguno de los dos casos anteriores, se informará con este estado cuando encuentre para ese mes una incidencia en el area de Facturacion.
                        strSql = " select area, incidencia, estado_incidencia, fecha_apertura, prioridad_negocio, titulo, e_s_estado,Reincidente from Relacion_INC_CUPS"
                                 + " where cups='" + r["cups20"].ToString() + "' and Mes_pendiente='" + aniomes + "'"
                                 + " order by Fecha_apertura asc";


                        areaincidencia = "";
                        incidencia = "";
                        estado_incidencia = "";
                        fecha_apertura = "";
                        prioridad_negocio = "";
                        titulo = "";
                        e_s_estado = "";
                        subestado_incidencia = "";
                        reincidente = "";
                        ExisteIncidencia = false;

                        dbIncidencia = new MySQLDB(MySQLDB.Esquemas.MED);
                        command = new MySqlCommand(strSql, dbIncidencia.con);
                        inci = command.ExecuteReader();

                        while (inci.Read())
                        {
                            ExisteIncidencia = true;
                            if (areaincidencia == "")
                            {
                                areaincidencia = inci["area"].ToString();

                                if (areaincidencia == "Contratacion")
                                    subestado_incidencia = "01.B09 Incidencia_Contratacion";
                                if (areaincidencia == "Medida")
                                    subestado_incidencia = "01.B10 Incidencia_Medida";
                                if (areaincidencia == "Facturacion")
                                    subestado_incidencia = "01.B11 Incidencia_Facturacion";

                                incidencia = inci["incidencia"].ToString();
                                estado_incidencia = inci["estado_incidencia"].ToString();
                                fecha_apertura = inci["fecha_apertura"].ToString();
                                prioridad_negocio = inci["prioridad_negocio"].ToString();
                                titulo = inci["titulo"].ToString();
                                e_s_estado = inci["e_s_estado"].ToString();
                                reincidente = inci["Reincidente"].ToString();
                            }
                            else
                            {
                                if (areaincidencia != inci["area"].ToString())
                                {
                                    areaincidencia = inci["area"].ToString();

                                    if (areaincidencia == "Contratacion")
                                        subestado_incidencia = subestado_incidencia + " - 01.B09 Incidencia_Contratacion";
                                    if (areaincidencia == "Medida")
                                        subestado_incidencia = subestado_incidencia + " - 01.B10 Incidencia_Medida";
                                    if (areaincidencia == "Facturacion")
                                        subestado_incidencia = subestado_incidencia + " - 01.B11 Incidencia_Facturacion";

                                    incidencia = incidencia + " - " + inci["incidencia"].ToString();
                                    estado_incidencia = estado_incidencia + " - " + inci["estado_incidencia"].ToString();
                                    fecha_apertura = fecha_apertura + " - " + inci["fecha_apertura"].ToString();
                                    prioridad_negocio = prioridad_negocio + " - " + inci["prioridad_negocio"].ToString();
                                    titulo = titulo + " - " + inci["titulo"].ToString();
                                    e_s_estado = e_s_estado + " - " + inci["e_s_estado"].ToString();
                                }
                            }
                        } // Fin  while (inci.Read())

                        dbIncidencia.CloseConnection();
                        if (ExisteIncidencia == true)
                        {
                            c--;
                            workSheet.Cells[f, c].Value = subestado_incidencia; c++;
                            workSheet.Cells[f, c].Value = IncidenciaFacturacion; c++;
                            workSheet.Cells[f, c].Value = Estado_FAC_SE; c++;
                            workSheet.Cells[f, c].Value = Titulo_FAC; c++;

                            workSheet.Cells[f, c].Value = incidencia; c++;
                            //REINCIDENTE
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = estado_incidencia; c++;
                            workSheet.Cells[f, c].Value = fecha_apertura; c++;
                            workSheet.Cells[f, c].Value = prioridad_negocio; c++;
                            workSheet.Cells[f, c].Value = titulo; c++;
                            workSheet.Cells[f, c].Value = e_s_estado; c++;
                        }
                        else
                        {
                         
                            workSheet.Cells[f, c].Value = IncidenciaFacturacion; c++;
                            workSheet.Cells[f, c].Value = Estado_FAC_SE; c++;
                            workSheet.Cells[f, c].Value = Titulo_FAC; c++;
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = ""; c++;
                        }
                        /////////////////////////////////////////////////////////////////////
                 
                        //FH_ALTA_SALESFORCE
                        ////workSheet.Cells[f, c].Value = ""; c++;

                        if (r["fh_alta_crto"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_alta_crto"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        else
                        {

                        }
                        c++;
                        // FIN FH_ALTA_SALESFORCE


                        //FECHA_ALTA_KEE*****************************************************
                        listaAltaKEE = pendienteWeb_B2B.GetCupsFechaAltaKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fec_act"]));

                        if (listaAltaKEE.Count == 0)
                        {
                            workSheet.Cells[f, c].Value = ""; c++;
                        }
                        else
                        {
                            if (listaAltaKEE[0] == "01/01/0001 0:00:00")
                            {
                                workSheet.Cells[f, c].Value = ""; c++;
                            }
                            else
                            {

                                foreach (var mm in listaAltaKEE)
                                {
                                    workSheet.Cells[f, c].Value = mm; c++;
                                    break;
                                }
                                //////workSheet.Cells[f, c].Value = lista; c++;
                            }
                        }

                        //FIN FH_ALTA_KEE*****************************************************

                        //FH_ALTA_SAP (de momento lo pongo en blanco y me quedo con la posición para rellenar más adelante
                        workSheet.Cells[f, c].Value = "";
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        FhAltaSap = c;
                        c++;

                        //FH_BAJA_SALESFORCE
                        if (r["fh_baja_crto"] != System.DBNull.Value) 
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_baja_crto"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        else
                        {

                        }
                        c++;
                        // FIN  FH_BAJA_SALESFORCE

                        //Paco, buscamos la fecha de Baja de SAP ******************************************
                        CadenaAuxiliar = r["id_crto_ext"].ToString();
                        CadenaAuxiliar = CadenaAuxiliar.PadLeft(12, '0'); //Relleno con ceros a la izquierda hasta 12

                        fhBajaSap = FechaBajaSap.GetCupsFechaBajaSAP(CadenaAuxiliar);

                        if (fhBajaSap.Count == 0)
                        {
                            workSheet.Cells[f, c].Value = ""; c++;
                        }
                        else
                        {
                                foreach (var mm in fhBajaSap)
                                {
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    workSheet.Cells[f, c].Value = mm; c++; 
                                    break;
                                }
                        }


                        //////strSql = "SELECT  max(cd_sec_crto), fh_baja from ed_owner.t_ed_h_sap_crto_front "
                        //////+ " WHERE id_crto_ext = '" + CadenaAuxiliar + "'"
                        //////+ " group by  fh_baja"
                        //////+ " order by  max(cd_sec_crto) desc";

                        //////dbRS = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                        //////commandRS = new OdbcCommand(strSql, dbRS.con);
                        //////rRS = commandRS.ExecuteReader();
                        //////while (rRS.Read())
                        //////{
                        //////    if (rRS["fh_baja"] != System.DBNull.Value)
                        //////        workSheet.Cells[f, c].Value = Convert.ToDateTime(rRS["fh_baja"]).Date;
                        //////    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        //////    break;
                        //////}
                        //////dbRS.CloseConnection();
                        //////c++;



                        //Fin - Fh_baja de SAP ***************************************************

                        if (r["cd_empr"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["cd_empr"].ToString();
                            Existetedhps = true;
                        }
                        else
                        {
                            Existetedhps = false;

                            strSql = "SELECT cd_empr,  cd_nif_cif_cli, de_tp_cli, tx_apell_cli, fh_alta_crto, fh_inicio_vers_crto"
                            + " ,ps.cd_tarifa_c, ps.cd_crto_comercial, ps.de_empr_distdora_nombre"
                            + " FROM cont.t_ed_h_ps_hist ps"
                            + " WHERE cups20 = '" + r["cups20"].ToString() + "'"
                            + " AND created_date = (SELECT MAX(created_date) FROM cont.t_ed_h_ps_hist"
                            + " WHERE cups20 = '" + r["cups20"].ToString() + "')";

                            dbAux = new MySQLDB(MySQLDB.Esquemas.GBL);
                            commandAux = new MySqlCommand(strSql, dbAux.con);
                            rAux = commandAux.ExecuteReader();
                            while (rAux.Read())
                            {
                                if (rAux["cd_empr"] != System.DBNull.Value)
                                    empresa = rAux["cd_empr"].ToString();
                                if (rAux["de_tp_cli"] != System.DBNull.Value)
                                    cliente = rAux["de_tp_cli"].ToString();
                                if (rAux["cd_nif_cif_cli"] != System.DBNull.Value)
                                    nif = rAux["cd_nif_cif_cli"].ToString();
                                if (rAux["tx_apell_cli"] != System.DBNull.Value)
                                    apellido = rAux["tx_apell_cli"].ToString();
                                if (rAux["fh_alta_crto"] != System.DBNull.Value)
                                    FechaAlta = Convert.ToDateTime(rAux["fh_alta_crto"]).Date;
                                if (rAux["fh_inicio_vers_crto"] != System.DBNull.Value)
                                    FechaInicio = Convert.ToDateTime(rAux["fh_inicio_vers_crto"]).Date;
                                if (rAux["cd_tarifa_c"] != System.DBNull.Value)
                                    Tarifa = rAux["cd_tarifa_c"].ToString();
                                if (rAux["cd_crto_comercial"] != System.DBNull.Value)
                                    contrato = rAux["cd_crto_comercial"].ToString();
                                if (rAux["de_empr_distdora_nombre"] != System.DBNull.Value)
                                    Distribuidora = rAux["de_empr_distdora_nombre"].ToString();
                            }
                            dbAux.CloseConnection();
                            workSheet.Cells[f, c].Value = empresa;
                        }
                        c++;

                        if (Existetedhps == false) //TIPO_CLIENTE
                        {
                            workSheet.Cells[f, c].Value = cliente;
                        }
                        else
                        {
                            if (r["de_tp_cli"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["de_tp_cli"].ToString();
                        }
                        c++;

                        if (Existetedhps == false) //NIF
                        {
                            workSheet.Cells[f, c].Value = nif;
                        }
                        else
                        {
                            if (r["cd_nif_cif_cli"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["cd_nif_cif_cli"].ToString();
                        }
                        c++;

                        if (Existetedhps == false) //Es la fh_alta_sap que me he guardado la posición, se pinta más arriba
                        {
                            if (FechaAlta != null)
                            {
                                workSheet.Cells[f, FhAltaSap].Value = Convert.ToDateTime(FechaAlta).Date;
                                workSheet.Cells[f, FhAltaSap].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                        }
                        else
                        {
                            if (r["fh_alta_crto"] != System.DBNull.Value)
                            {
                                workSheet.Cells[f, FhAltaSap].Value = Convert.ToDateTime(r["fh_alta_crto"]).Date;
                                workSheet.Cells[f, FhAltaSap].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                        }

                        //c++;

                        if (Existetedhps == false)
                        {
                            workSheet.Cells[f, intCliente].Value = apellido;
                        }
                        else
                        {
                            if (r["tx_apell_cli"] != System.DBNull.Value)
                                workSheet.Cells[f, intCliente].Value = r["tx_apell_cli"].ToString();
                        }
                        //c++;


                        if (Existetedhps == false) //FPSERCON
                        {
                            if (FechaInicio != null)
                            {
                                workSheet.Cells[f, c].Value = Convert.ToDateTime(FechaInicio).Date;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                        }
                        else
                        {
                            if (r["fh_inicio_vers_crto"] != System.DBNull.Value)
                            {
                                workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_inicio_vers_crto"]).Date;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                        }
                        c++;

                        if (r["id_instalacion"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["id_instalacion"].ToString();
                        c++;

                        if (Existetedhps == false)
                        {
                            workSheet.Cells[f, c].Value = Tarifa;
                        }
                        else
                        {
                            if (r["cd_tarifa_c"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["cd_tarifa_c"].ToString();
                        }
                        c++;

                        if (Existetedhps == false)
                        {
                            workSheet.Cells[f, c].Value = contrato;
                        }
                        else
                        {
                            if (r["cd_crto_comercial"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["cd_crto_comercial"].ToString();
                        }
                        c++;
                        if (Existetedhps == false)
                        {
                            workSheet.Cells[f, c].Value = Distribuidora;
                        }
                        else
                        {
                            if (r["de_empr_distdora_nombre"] != System.DBNull.Value)
                                workSheet.Cells[f, c].Value = r["de_empr_distdora_nombre"].ToString();
                        }
                        c++;

                        if (r["de_estado"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["de_estado"].ToString();
                        c++;

                        if (r["de_subestado"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["de_subestado"].ToString();
                        c++;

                        ///PACO PRUEBA
                        //////if (r["cups20"] != System.DBNull.Value)
                        //////{
                        //////    workSheet.Cells[f, c].Value = GetDiasEstado(r["cups20"].ToString());
                        //////    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //////}
                        //////c++;

                        if (r["TAM"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        }
                        c++;

                        System.Diagnostics.Debug.WriteLine(r["cups20"].ToString());
                        // Paco - añadimos la ultima fecha desde y hasta facturada t_ed_h_sap_facts
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                        List<string> listaInicio = new List<string>();
                        listaInicio = pendienteSAPKEE.GetCupsInicioUltimaFacturada(r["cups20"].ToString());
                        if (listaInicio != null)
                        {
                            if (listaInicio.Count == 0)
                            {
                                workSheet.Cells[f, c].Value = ""; c++;
                            }
                            else
                            {
                                if (listaInicio[0] == "01/01/0001 0:00:00")
                                {
                                    workSheet.Cells[f, c].Value = ""; c++;
                                }
                                else
                                {

                                    foreach (var mm in listaInicio)
                                    {
                                        workSheet.Cells[f, c].Value = mm; c++;
                                        break;
                                    }
                                    ////workSheet.Cells[f, c].Value = lista; c++;
                                }
                            }
                        }
                        else {
                            workSheet.Cells[f, c].Value = ""; c++;
                        }

                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                        List<string> listaFin = new List<string>();
                        listaFin = pendienteSAPKEE.GetCupsFinalUltimaFacturada(r["cups20"].ToString());

                        if (listaFin != null)
                        {
                            if (listaFin.Count == 0 )
                            {
                                workSheet.Cells[f, c].Value = ""; c++;
                            }
                            else
                            {
                                if (listaFin[0] == "01/01/0001 0:00:00")
                                {
                                    workSheet.Cells[f, c].Value = ""; c++;
                                }
                                else
                                {

                                    foreach (var mm in listaFin)
                                    {
                                        workSheet.Cells[f, c].Value = mm; c++;
                                        break;
                                    }
                                    //////workSheet.Cells[f, c].Value = lista; c++;
                                }
                            }
                        }
                        else {
                            workSheet.Cells[f, c].Value = ""; c++;
                        }

                        System.Diagnostics.Debug.WriteLine(f);

                        subestadosKronos.existe = false;

                        if (subestados_sap.AreaResponsableMedida(r["cd_subestado"].ToString()))
                        {

                            subestadosKronos.existe = false;

                            //Paco -- He duplicado funciones para marcar los casos que no hay datos en kronos para la fecha GetEstadoKEEDetalle y GetCupsDetalle
                            //Faltaria controlar los casos en los que vienen segmentos, es decir, para el mes que se mira de Sap vienen dos segmentos
                            //en Kronos, habria que quedarse con el de fecha más antigua, he puesto codigo sin probar para este caso en GetCupsDetalle
                            //Lo que hago es crear a pelo un estado "Discrepancia: ..... " que es la etiqueta que pongo y que no está en la tabla parametrica estados_kee_param 

                            if (multipunto == "N")
                            {
                                subestadosKronos.GetEstadoKEEDetalle(pendienteWeb_B2B.GetCupsDetalle(r["cups20"].ToString(),
                                    Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]), Convert.ToDateTime(r["fec_act"])));
                            }
                            else
                            {
                                subestadosKronos.GetEstadoKEEDetalleMultipunto(pendienteWeb_B2B.GetCupsDetalle(r["cups20"].ToString(),
                                    Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]), Convert.ToDateTime(r["fec_act"])));
                            }

                            //////subestadosKronos.GetEstadoKEEDetalle(pendienteWeb_B2B.GetCupsDetalle(r["cups20"].ToString(),
                            //////    Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]), Convert.ToDateTime(r["fec_act"])));

                            if (subestadosKronos.existe)
                            {
                                BuscaCadena = subestadosKronos.descripcion_estado;
                                List<string> lista = new List<string>();

                                if (BuscaCadena.IndexOf("Discrepancia") >= 0)
                                {
                                    if (BuscaCadena.IndexOf("Periodo SAP encontrado parcialmente en el informe pendiente de KEE") >= 0 )
                                    {
                                        workSheet.Cells[f, c].Value = subestadosKronos.estado_periodo; c++;
                                        workSheet.Cells[f, c].Value = subestadosKronos.area_responsable; c++;
                                        workSheet.Cells[f, c].Value = "01.B05 Error Sistemas KEE - SAP - Incoherencia periodos KEE-SAP"; c++;
                                        // Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el estado de KEE lo he guardado en descripcion_subestado
                                        //en lugar de en descripcion_estado
                                        workSheet.Cells[f, c].Value = subestadosKronos.descripcion_subestado; c++;
                                        string[] cadena = subestadosKronos.descripcion_estado.Split(';');
                                        string[] cadenaAux = cadena[2].Split('/');
                                        workSheet.Cells[f, c].Value = cadenaAux[0].Substring(1, 8).Substring(6, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(4, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(0, 4); c++;
                                        workSheet.Cells[f, c].Value = cadenaAux[1].Substring(0, 8).Substring(6, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(4, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(0, 4); c++;
                                    }
                                    //|| BuscaCadena.IndexOf("Periodos del pendiente de KEE contenidos en las fechas de SAP") >= 0 || BuscaCadena.IndexOf("Periodo SAP contenido dentro de un informe pendiente de KEE") >= 0
                                    else
                                    {
                                        if (BuscaCadena.IndexOf("Periodos del pendiente de KEE contenidos en las fechas de SAP") >= 0 || BuscaCadena.IndexOf("Periodo SAP contenido dentro de un informe pendiente de KEE") >= 0)
                                        {
                                            workSheet.Cells[f, c].Value = subestadosKronos.estado_periodo; c++;
                                            workSheet.Cells[f, c].Value = subestadosKronos.area_responsable; c++;
                                            // Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el subestado de KEE lo he guardado en temporal de momento
                                            workSheet.Cells[f, c].Value = subestadosKronos.temporal; c++;
                                            // Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el estado de KEE lo he guardado en descripcion_subestado
                                            //en lugar de en descripcion_estado
                                            workSheet.Cells[f, c].Value = subestadosKronos.descripcion_subestado; c++;
                                            string[] cadena = subestadosKronos.descripcion_estado.Split(';');
                                            string[] cadenaAux = cadena[2].Split('/');
                                            workSheet.Cells[f, c].Value = cadenaAux[0].Substring(1, 8).Substring(6, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(4, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(0, 4); c++;
                                            workSheet.Cells[f, c].Value = cadenaAux[1].Substring(0, 8).Substring(6, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(4, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(0, 4); c++;
                                        }
                                        else
                                        {
                                            //si no está el CUPS en la tabla "ed_owner.T_ED_H_GEST_DIAR_PS_B2B" he obtenido: "Discrepancia: No existe el cups en el informe del pendiente de KEE"
                                            // el cups está en la tabla, pero dentro de los periodos que hay no no hay ningún periodo que coincida ni total ni parcialmente
                                            //      Marcado: "Discrepancia: No existen fechas en el informe pendiente de KEE para el periodo;(" + Convert.ToString(fecha_desde.ToString("yyyyMMdd")) + "/" + Convert.ToString(fecha_hasta.ToString("yyyyMMdd")) + ")")

                                            if (BuscaCadena.IndexOf("No existe el cups en el informe del pendiente de KEE") >= 0)
                                            {
                                                workSheet.Cells[f, c].Value = ""; c++;
                                                workSheet.Cells[f, c].Value = ""; c++;
                                                workSheet.Cells[f, c].Value = "01.B07 Error Sistemas KEE-SAP - Pdte SAP-No existe en KEE"; c++;
                                                workSheet.Cells[f, c].Value = ""; c++;
                                                workSheet.Cells[f, c].Value = ""; c++;
                                                workSheet.Cells[f, c].Value = ""; c++;

                                            }
                                            else {

                                                if (BuscaCadena.IndexOf("No existen fechas en el informe pendiente de KEE para el periodo") >= 0)
                                                {
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = "01.B06 Error Sistemas KEE-SAP - Pdte SAP-No pdte KEE"; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                }
                                                else
                                                {
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                }

                                            } //Fin  if (BuscaCadena.IndexOf("No existe el cups en el informe del pendiente de KEE") >= 0)

                                        } // fin  if (BuscaCadena.IndexOf("Periodos del pendiente de KEE contenidos en las fechas de SAP") >= 0 || BuscaCadena.IndexOf("Periodo SAP contenido dentro de un informe pendiente de KEE") >= 0)

                                    } // Fin  if (BuscaCadena.IndexOf("Periodo SAP encontrado parcialmente en el informe pendiente de KEE") >= 0 )

                                    ///lista = pendienteWeb_B2B.GetCupsFinKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]));
                                    lista = pendienteWeb_B2B.GetCupsFinKEENew(r["cups20"].ToString(), Convert.ToDateTime(r["fec_act"]));

                                    if (lista.Count == 0)
                                    {
                                        workSheet.Cells[f, c].Value = ""; c++;
                                    }
                                    else
                                    {
                                        if (lista[0] == "01/01/0001 0:00:00")
                                        {
                                            workSheet.Cells[f, c].Value = ""; c++;
                                        }
                                        else
                                        {

                                            foreach (var mm in lista)
                                            {
                                                workSheet.Cells[f, c].Value = mm; c++;
                                                break;
                                            }
                                            //////workSheet.Cells[f, c].Value = lista; c++;
                                        }
                                    }
                                    workSheet.Cells[f, c].Value = subestadosKronos.descripcion_estado; c++;
                                    //Subestado Global
                                    if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                    {
                                        workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoKEE].Value;c++;
                                    }
                                    else {
                                        workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoSAP].Value; c++;
                                    }
                                }
                                else
                                {
                                    workSheet.Cells[f, c].Value = subestadosKronos.estado_periodo; c++;
                                    workSheet.Cells[f, c].Value = subestadosKronos.area_responsable; c++;
                                    workSheet.Cells[f, c].Value = subestadosKronos.descripcion_subestado; c++;
                                    workSheet.Cells[f, c].Value = subestadosKronos.descripcion_estado; c++;
                                    //Paco 22/03/2024 fecha_desde_kee y fecha_hasta_kee
                                    if (r["fh_desde"] != System.DBNull.Value)
                                    {
                                        workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_desde"]).Date;
                                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    }
                                    c++;

                                    if (r["fh_hasta"] != System.DBNull.Value)
                                    {
                                        workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_hasta"]).Date;
                                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    }
                                    c++;
                                    ///lista = pendienteWeb_B2B.GetCupsFinKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]));
                                    lista = pendienteWeb_B2B.GetCupsFinKEENew(r["cups20"].ToString(), Convert.ToDateTime(r["fec_act"]));
                                    if (lista.Count == 0)
                                    {
                                        workSheet.Cells[f, c].Value = ""; c++;
                                    }
                                    else
                                    {
                                        if (lista[0] == "01/01/0001 0:00:00")
                                        {
                                            workSheet.Cells[f, c].Value = ""; c++;
                                        }
                                        else
                                        {

                                            foreach (var mm in lista)
                                            {
                                                workSheet.Cells[f, c].Value = mm; c++;
                                                break;
                                            }
                                            //////workSheet.Cells[f, c].Value = lista; c++;
                                        }
                                    }

                                    //////workSheet.Cells[f, c].Value = ""; c++;
                                    //////workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;

                                    //Subestado Global
                                    if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                    {
                                        workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoKEE].Value; c++;
                                    }
                                    else
                                    {
                                        workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoSAP].Value; c++;
                                    }
                                    //Fin Subestado Global
                                }

                                //Paco 21/05/2024 Calculo del numero dias KEE en Kronos ********************
                                if (blnLanzarDiasKEE == true)
                                {
                                    //////DiasExistenciaKEE = "";
                                    //Primero, busco si ya está el registro en la tabla
                                    //////if (BuscaCadena.IndexOf("Discrepancia") >= 0)
                                    //////{
                                    //////    DiasExistenciaKEE = FuncionDiasKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), subestadosKronos.descripcion_subestado);
                                    //////}
                                    //////else
                                    //////{
                                    //////    DiasExistenciaKEE = FuncionDiasKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), subestadosKronos.descripcion_estado);
                                    DiasExistenciaKEE = FuncionDiasKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), subestadosKronos.estado_periodo);
                                    //////}
                                }
                                else
                                {
                                    /// Para que recupere los dias en ejecuciones sucesivas el mismo día.
                                    //////if (BuscaCadena.IndexOf("Discrepancia") >= 0)
                                    //////{
                                    //////    DiasExistenciaKEE = FuncionDiasKEERecuperarDiasTabla(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), subestadosKronos.descripcion_subestado);
                                    //////}
                                    //////else
                                    //////{
                                    //////    DiasExistenciaKEE = FuncionDiasKEERecuperarDiasTabla(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), subestadosKronos.descripcion_estado);
                                    //////}    
                                    DiasExistenciaKEE = FuncionDiasKEERecuperarDiasTabla(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), subestadosKronos.estado_periodo);
                                }

                                //Fin Calculo del numero dias KEE en Kronos  ***************
                            }
                            else
                            {
                                workSheet.Cells[f, c].Value = ""; c++;
                                workSheet.Cells[f, c].Value = ""; c++;
                                workSheet.Cells[f, c].Value = ""; c++;
                                workSheet.Cells[f, c].Value = ""; c++;
                                workSheet.Cells[f, c].Value = ""; c++;
                                workSheet.Cells[f, c].Value = ""; c++;
                                workSheet.Cells[f, c].Value = ""; c++;
                                workSheet.Cells[f, c].Value = ""; c++;
                                //Subestado Global
                                if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                {
                                    workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoKEE].Value; c++;
                                }
                                else
                                {
                                    workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoSAP].Value; c++;
                                }
                                //Fin Subestado Global
                            }

                        }
                        else
                        {
                            //System.Diagnostics.Debug.WriteLine(r["cups20"].ToString() + "-" + r["de_subestado"].ToString() + "-" + subestadosKronos.descripcion_subestado + "-" + subestadosKronos.descripcion_estado);
                            //Para que llegue hasta la última columna aunque no exista en Kronos
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = ""; c++;
                            // Fecha fin KEE (se pinta aunque no sea uno de los estados marcados como KEE
                            // workSheet.Cells[f, c].Value = ""; c++;
                            List<string> listaAux = new List<string>();
                            listaAux = pendienteWeb_B2B.GetCupsFinKEENew(r["cups20"].ToString(), Convert.ToDateTime(r["fec_act"]));
                            if (listaAux.Count == 0)
                            {
                                workSheet.Cells[f, c].Value = ""; c++;
                            }
                            else
                            {
                                if (listaAux[0] == "01/01/0001 0:00:00")
                                {
                                    workSheet.Cells[f, c].Value = ""; c++;
                                }
                                else
                                {

                                    foreach (var mm in listaAux)
                                    {
                                        workSheet.Cells[f, c].Value = mm; c++;
                                        break;
                                    }
                                    //////workSheet.Cells[f, c].Value = lista; c++;
                                }
                            }
                            workSheet.Cells[f, c].Value = ""; c++;

                            //Subestado Global
                            if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                            {
                                workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoKEE].Value; c++;  
                            }
                            else
                            {
                                workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoSAP].Value; c++; 
                            } 

                        } // Fin  if (subestados_sap.AreaResponsableMedida(r["cd_subestado"].ToString()))



                        //Calculo DIAS ESTADO GLOBAL *************************
                        if (blnLanzarDiasKEE == true) // Sólo se actualizan los datos en la tabla la primera vez, por si hay ejecuciones sucesivas
                        {
                            if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                            {
                                DiasExistenciaKEEGlobal = FuncionDiasKEEGlobal(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), workSheet.Cells[f, PosicionSubEstadoKEE].Value.ToString());
                            }
                            else
                            {
                                DiasExistenciaKEEGlobal = FuncionDiasKEEGlobal(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), workSheet.Cells[f, PosicionSubEstadoSAP].Value.ToString());
                            }
                        }
                        else
                        {
                            if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                            {
                                DiasExistenciaKEEGlobal = FuncionDiasKEERecuperarDiasTablaGlobal(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), workSheet.Cells[f, PosicionSubEstadoKEE].Value.ToString());
                            }
                            else
                            {
                                DiasExistenciaKEEGlobal = FuncionDiasKEERecuperarDiasTablaGlobal(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), workSheet.Cells[f, PosicionSubEstadoSAP].Value.ToString());
                            }
                        } 
                        //FIN DIAS ESTADO GLOBAL*******************************

                       
                        if (areaincidencia != "" & subestadosKronos.existe == true)
                        {
                            Auxiliar = Reincidente(areaincidencia, subestadosKronos.area_responsable);
                            if (Auxiliar != "")
                            {
                                ////c--;
                                workSheet.Cells[f, PosicionSubEstadoGlobal].Value = Auxiliar;
                                workSheet.Cells[f, PosicionReincidente].Value = reincidente;
                            }
                        }
                        //Fin Paco 09/05/2024***********************************************************************************************

                        ///14/05/2024 New EStados Global***********************************************

                        Auxiliar = "";
                        Auxiliar = EstadosAreas(workSheet.Cells[f, PosicionSubEstadoGlobal].Value.ToString(), cups20, aniomes.ToString(), 1,0, false);

                        if (Auxiliar != "")
                        {
                            string[] CadenaEstados = Auxiliar.Split(';');
                            workSheet.Cells[f, c].Value = CadenaEstados[0]; c++;
                            workSheet.Cells[f, c].Value = CadenaEstados[1]; c++;
                            workSheet.Cells[f, c].Value = CadenaEstados[2]; c++;
                        }
                        else
                        {
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = ""; c++;
                            workSheet.Cells[f, c].Value = ""; c++;

                            if (SubestadosNoEncontrados == "")
                            {
                                SubestadosNoEncontrados = "Hoja ES: "+ cups20 + " - " + Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd") + " - " + Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd") + " - " + r["de_estado"].ToString() + " -" + r["de_subestado"].ToString() + "- No existe el subestado en la tabla paramétrica";
                             
                            }
                            else
                            {
                                SubestadosNoEncontrados = SubestadosNoEncontrados + ";Hoja ES: " + cups20 + " - " + Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd") + " - " + Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd") + " - " + r["de_estado"].ToString() + " -" + r["de_subestado"].ToString() + "- No existe el subestado en la tabla paramétrica";
                            }
                        }

                        ////////Fin New EStados Global*******************************************************

                        //// Tratamiento multipuntos
                        workSheet.Cells[f, c].Value =multipunto; c++;

                        // "Nº DÍAS ESTADO KEE"
                        if (subestadosKronos.existe == true)
                        {
                            workSheet.Cells[f, PosicionDiasEstadoKEE].Value = DiasExistenciaKEE; c++;
                        }

                        // Pego Nº DIAS ESTADO GLOBAL, va a continuación de DIAS ESTADO KEE
                        workSheet.Cells[f, PosicionDiasEstadoKEE+1].Value = DiasExistenciaKEEGlobal; c++;


                    } //Fin bucle principal     while (r.Read())


                    intConteoESDetalle = f-1;
                    db.CloseConnection();

                    headerCells = workSheet.Cells[1, 1, 1, c];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;

                    //08/02/2024 Paco, mover columna de fecha baja de Kronos despues de fecha de baja de SAP
                    workSheet.InsertColumn(PosicionBajaSap, 1); //Inserta en la posicion 12 una columna en blanco(detrás de FH_BAJA_SAP)
                    workSheet.Cells[1, PosicionBajaKEE, f, PosicionBajaKEE].Copy(workSheet.Cells[1, PosicionBajaSap, f, PosicionBajaSap]); //Copia la columna 31 (Fecha BAJA KEE) en la posición creada anteriormente
                    workSheet.DeleteColumn(PosicionBajaKEE); //Borra la columna 31 Fecha BAJA KEE

                    //Pegamos AREA, Subestadoglobal, estado_global y estado global a reportar despues de FH_HASTA *** (!!!OJO, he metido ya por medio FECHA BAJA KEE (Sumo una más)!!)****

                    //Pegamos Subestadoglobal, estado_global y estado global a reportar despues de FH_HASTA *** (!!!OJO, he metido ya por medio FECHA BAJA KEE (Sumo una más)!!)****
                    workSheet.InsertColumn(PosicionPegarSubEstadoGlobal, 1); //Inserta en la posicion 5 una columna en blanco(detrás de FH_HASTA)
                    workSheet.Cells[1, PosicionSubEstadoGlobal + 1, f, PosicionSubEstadoGlobal + 1].Copy(workSheet.Cells[1, PosicionPegarSubEstadoGlobal, f, PosicionPegarSubEstadoGlobal]);

                    workSheet.InsertColumn(PosicionPegarSubEstadoGlobal + 1, 1); //inserto en posición 6 (ya he pegado subestadoglobal)
                    workSheet.Cells[1, PosicionSubEstadoGlobal + 3, f, PosicionSubEstadoGlobal + 3].Copy(workSheet.Cells[1, PosicionPegarSubEstadoGlobal + 1, f, PosicionPegarSubEstadoGlobal + 1]);

                    workSheet.InsertColumn(PosicionPegarSubEstadoGlobal + 2, 1); //inserto en posición 7 (ya he pegado subestadoglobal)
                    workSheet.Cells[1, PosicionSubEstadoGlobal + 5, f, PosicionSubEstadoGlobal + 5].Copy(workSheet.Cells[1, PosicionPegarSubEstadoGlobal + 2, f, PosicionPegarSubEstadoGlobal + 2]);

                    // borro las originales (tres a la derecha, he insertado cuatro columnas)
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal+3);
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal+3);
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal+3);
                   

                    //Pongo multipunto  antes de discrepancias
                    //////workSheet.Cells[f, c].Value = "Fecha BAJA KEE"; c++;
                    //////PosicionBajaKEE = c;
                    //////workSheet.Cells[f, c].Value = "Discrepancias"; c++;
                    //////PosicionSubEstadoGlobal = c;
                    //////workSheet.Cells[f, c].Value = "Subestado global"; c++;
                    //////workSheet.Cells[f, c].Value = "ESTADO GLOBAL"; c++;
                    //////workSheet.Cells[f, c].Value = "ESTADO GLOBAL A REPORTAR"; c++;
                    //////workSheet.Cells[f, c].Value = "multipunto";
                    
                    // He quitado  lo que hay entre discrepancias y multipuntos, las posiciones me valen, ahora he metido los globales por delante
                    workSheet.InsertColumn(PosicionBajaKEE + 3, 1);
                    workSheet.Cells[1, PosicionSubEstadoGlobal + 4, f, PosicionSubEstadoGlobal + 4].Copy(workSheet.Cells[1, PosicionBajaKEE + 3, f, PosicionBajaKEE + 3]); //Copia la columna 31 (Fecha BAJA KEE) en la posición creada anteriormente
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal + 4);

                    //workSheet.Cells[1,5,100,5].Copy(workSheet.Cells[1,2,100,2]);
                    //Copia la columna 5 en la columna 2 Básicamente Source.Copy(Destino) Esto solo copiaría las primeras 100 filas.


                    //La columna area la pongo detrás de fh_hasta
                    workSheet.InsertColumn(PosicionPegarSubEstadoGlobal, 1);
                    // Ya he quitado lo que originalmente estaba detras de posiciónSubEstado Global, sólo queda el area en esa poscion
                    workSheet.Cells[1, PosicionSubEstadoGlobal +3, f, PosicionSubEstadoGlobal+3 ].Copy(workSheet.Cells[1, PosicionPegarSubEstadoGlobal, f, PosicionPegarSubEstadoGlobal]);
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal + 3);

                    //El final es discrepancias seguido de multipuntos, los tengo que cambiar de orden --> discrepancia=48 y multipunto=49 después de las inserciones de columna
                    //no lo puedo hacer durante el proceso y poner multipuntos antes, ya que las discrepancias es un proceso aparte y separado
                    //Busco donde ha quedado discrepancias
                    intColumna = 1;
                    while (workSheet.Cells[1, intColumna].Value != "" & workSheet.Cells[1, intColumna].Value != "Discrepancias")
                    {
                        intColumna = intColumna + 1;
                    }
                    workSheet.InsertColumn(intColumna, 1); //Inserta  columna en blanco delante de discrepancias
                    workSheet.Cells[1, intColumna+2, f, intColumna + 2].Copy(workSheet.Cells[1, intColumna, f, intColumna]); //copio multipunto en la columna que he insertado
                    workSheet.DeleteColumn(intColumna + 2); //Borra la columna original de multipunto

                    allCells = workSheet.Cells[1, 1, f, c+7]; //Pongo +7 para las columnas de incidencias
                    allCells.AutoFitColumns();

                    //workSheet.View.FreezePanes(2, 1);
                    //workSheet.Cells["A1:S1"].AutoFilter = true;
                    //allCells.AutoFitColumns();

                    f = 1;

                    #endregion

                    #region Detalle POR MT-BTE
                    f = 1;
                    c = 1;

                    workSheet = excelPackage.Workbook.Worksheets.Add("Detalle POR MT-BTE");

                    headerCells = workSheet.Cells[1, 1, 1, 32];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;
                    allCells = workSheet.Cells[1, 1, 50, 50];

                    //for (int i = 1; i <= c; i++)
                    //{
                    //    workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //    workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    //}


                    //Ponemos columnas directamente
                    f = 1;
                    c = 1;

                    workSheet.Cells[f, c].Value = "CUPS20"; c++;
                    workSheet.Cells[f, c].Value = "PERIODO"; c++;
                    workSheet.Cells[f, c].Value = "MES"; c++;
                    workSheet.Cells[f, c].Value = "FH_DESDE"; c++;
                    workSheet.Cells[f, c].Value = "FH_HASTA"; c++;
                    PosicionPegarSubEstadoGlobal = c;
                    workSheet.Cells[f, c].Value = "IMPORTE PDTE FACTURAR"; c++;
                    workSheet.Cells[f, c].Value = "MESES PDTES FACTURAR"; c++;
                    workSheet.Cells[f, c].Value = "ÁGORA"; c++;
                    workSheet.Cells[f, c].Value = "CLIENTE"; c++;
                    workSheet.Cells[f, c].Value = "DIAS_ESTADO_SAP"; c++;
                    PosicionDiasEstadoKEE = c;
                    workSheet.Cells[f, c].Value = "DIAS_ESTADO_KEE"; c++;
                    workSheet.Cells[f, c].Value = "DIAS_ESTADO_GLOBAL"; c++;
                    workSheet.Cells[f, c].Value = "Incidencia Facturacion"; c++;
                    workSheet.Cells[f, c].Value = "Estado_Fac_SE"; c++;
                    workSheet.Cells[f, c].Value = "Titulo_Fac"; c++;
                    workSheet.Cells[f, c].Value = "Incidencia Medida"; c++;
                    PosicionReincidente = c;
                    workSheet.Cells[f, c].Value = "Reincidente"; c++;
                    workSheet.Cells[f, c].Value = "Estado incidencia"; c++;
                    workSheet.Cells[f, c].Value = "Fecha_Apertura"; c++;
                    workSheet.Cells[f, c].Value = "Prioridad"; c++;
                    workSheet.Cells[f, c].Value = "Titulo"; c++;
                    workSheet.Cells[f, c].Value = "E_S_Estado"; c++;
                    workSheet.Cells[f, c].Value = "FH_ALTA_SALESFORCE"; c++;
                    workSheet.Cells[f, c].Value = "FH_ALTA_KEE"; c++;
                    workSheet.Cells[f, c].Value = "FH_ALTA_SAP"; c++;
                    workSheet.Cells[f, c].Value = "FH_BAJA_SALESFORCE"; c++;
                    PosicionBajaSap = c;
                    workSheet.Cells[f, c].Value = "FH_BAJA_SAP"; c++;
                    workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                    workSheet.Cells[f, c].Value = "SEGMENTO"; c++;
                    workSheet.Cells[f, c].Value = "TIPO CLIENTE"; c++;
                    workSheet.Cells[f, c].Value = "NIF"; c++;
                    workSheet.Cells[f, c].Value = "FPSERCON"; c++;
                    workSheet.Cells[f, c].Value = "Nº INSTALACIÓN"; c++;
                    workSheet.Cells[f, c].Value = "TARIFA"; c++;
                    workSheet.Cells[f, c].Value = "CONTRATO"; c++;
                    
                    workSheet.Cells[f, c].Value = "DISTRIBUIDORA"; c++;
                    workSheet.Cells[f, c].Value = "ESTADO"; c++;
                    PosicionSubEstadoSAP = c;
                    workSheet.Cells[f, c].Value = "SUBESTADO"; c++;
                    //////workSheet.Cells[f, c].Value = "DIAS_ESTADO"; c++;
                    workSheet.Cells[f, c].Value = "TAM"; c++;
                   
                    workSheet.Cells[f, c].Value = "ULT FH DESDE FACTURADA"; c++;
                    workSheet.Cells[f, c].Value = "ULT FH HASTA FACTURADA"; c++;

                    workSheet.Cells[f, c].Value = "Estado periodo KEE"; c++;
                    workSheet.Cells[f, c].Value = "Área responsable KEE"; c++;
                    PosicionSubEstadoKEE = c;
                    workSheet.Cells[f, c].Value = "Subestado KEE"; c++;
                    workSheet.Cells[f, c].Value = "Estado KEE"; c++;
                    workSheet.Cells[f, c].Value = "FH_DESDE_KEE"; c++;
                    workSheet.Cells[f, c].Value = "FH_HASTA_KEE"; c++;
                    workSheet.Cells[f, c].Value = "Fecha BAJA KEE"; c++;
                    PosicionBajaKEE = c;
                    workSheet.Cells[f, c].Value = "Discrepancias"; c++;
                    PosicionSubEstadoGlobal = c;
                    workSheet.Cells[f, c].Value = "Subestado global"; c++;
                    workSheet.Cells[f, c].Value = "ESTADO GLOBAL"; c++;
                    workSheet.Cells[f, c].Value = "ESTADO GLOBAL A REPORTAR"; c++;
                    workSheet.Cells[f, c].Value = "AREA ESTADO"; c++;
                    workSheet.Cells[f, c].Value = "multipunto";

                    strSql = DetalleExcel(fecha_informe, lista_empresas_PT, lista_segmentos_MT_BTE);

                    Segmentos = ",MT,BTE,AT,MAT,";

                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        ///07-11-2024
                        if (r["cd_segmento_ptg_Con_compor"] != System.DBNull.Value)
                        {
                            if (Segmentos.IndexOf(r["cd_segmento_ptg_Con_compor"].ToString()) > 0)
                            {
                        //Fin ///07-11-2024
                                f++;
                                c = 1;

                                empresa = ""; 
                                nif = "";
                                cliente = "";
                                apellido = "";
                                FechaAlta = null;
                                FechaInicio = null;
                                Tarifa = "";
                                contrato = "";
                                Distribuidora = "";
                                multipunto = "";
                                cups20 = "";

                                //Vemos si es o no multipunto
                                List<string> Multipuntos = new List<string>();
                                Multipuntos = pendienteWeb_B2B.GetCupsMultipunto(r["cups20"].ToString());
                                if (Multipuntos.Count > 0)
                                {
                                    if (Multipuntos[0] == "False")
                                    {
                                        multipunto = "N";
                                    }
                                    else
                                    {
                                        multipunto = "S";
                                    }

                                }
                                else
                                {
                                    multipunto = "N";
                                }

                                if (r["cups20"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = r["cups20"].ToString();
                                    cups20 = r["cups20"].ToString();
                                }   
                                c++;
                                ///PERIODO
                                workSheet.Cells[f, c].Value = ""; c++;

                                if (r["mes"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["mes"]);
                                    aniomes = Convert.ToInt32(r["mes"]);
                                    fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                        Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                                }
                                c++;

                                if (r["fh_desde"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_desde"]).Date;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (r["fh_hasta"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_hasta"]).Date;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                //Para controlar los blancos, miramos en t_ed_h_ps_hist
                                if (r["cd_empr"] != System.DBNull.Value)
                                {   
                                    Existetedhps = true;
                                }
                                else
                                {
                                    Existetedhps = false;

                                    strSql = "SELECT cd_empr,  cd_nif_cif_cli, de_tp_cli, tx_apell_cli, fh_alta_crto, fh_inicio_vers_crto"
                                    + " ,ps.cd_tarifa_c, ps.cd_crto_comercial, ps.de_empr_distdora_nombre"
                                    + " FROM cont.t_ed_h_ps_pt_hist ps"
                                    + " WHERE cups20 = '" + r["cups20"].ToString() + "'"
                                    + " AND created_date = (SELECT MAX(created_date) FROM cont.t_ed_h_ps_pt_hist"
                                    + " WHERE cups20 = '" + r["cups20"].ToString() + "')";

                                    dbAux = new MySQLDB(MySQLDB.Esquemas.GBL);
                                    commandAux = new MySqlCommand(strSql, dbAux.con);
                                    rAux = commandAux.ExecuteReader();
                                    while (rAux.Read())
                                    {
                                        if (rAux["cd_empr"] != System.DBNull.Value)
                                            empresa = rAux["cd_empr"].ToString();
                                        if (rAux["de_tp_cli"] != System.DBNull.Value)
                                            cliente = rAux["de_tp_cli"].ToString();
                                        if (rAux["cd_nif_cif_cli"] != System.DBNull.Value)
                                            nif = rAux["cd_nif_cif_cli"].ToString();
                                        if (rAux["tx_apell_cli"] != System.DBNull.Value)
                                            apellido = rAux["tx_apell_cli"].ToString();
                                        if (rAux["fh_alta_crto"] != System.DBNull.Value)
                                            FechaAlta = Convert.ToDateTime(rAux["fh_alta_crto"]).Date;
                                        if (rAux["fh_inicio_vers_crto"] != System.DBNull.Value)
                                            FechaInicio = Convert.ToDateTime(rAux["fh_inicio_vers_crto"]).Date;
                                        if (rAux["cd_tarifa_c"] != System.DBNull.Value)
                                            Tarifa = rAux["cd_tarifa_c"].ToString();
                                        if (rAux["cd_crto_comercial"] != System.DBNull.Value)
                                            contrato = rAux["cd_crto_comercial"].ToString();
                                        if (rAux["de_empr_distdora_nombre"] != System.DBNull.Value)
                                            Distribuidora = rAux["de_empr_distdora_nombre"].ToString();
                                    }
                                    dbAux.CloseConnection();
                                }
                                //// Fin mirar en el historico

                                //Importe PDTE facturar
                                if (r["mes"] != System.DBNull.Value && r["TAM"] != System.DBNull.Value)
                                {
                                    //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                                    meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["tam"]) * meses_pdtes;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                                }
                                else
                                {
                                    c++;
                                }

                       

                                if (r["mes"] != System.DBNull.Value)
                                {
                                    //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                                    meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                                    workSheet.Cells[f, c].Value = meses_pdtes;
                                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                }
                                c++;

                                if (meses_pdtes > 1)
                                {
                                    workSheet.Cells[f, 2].Value = ">1P";
                                }
                                else
                                {
                                    workSheet.Cells[f, 2].Value = "1P";
                                }

                                if (r["agora"] != System.DBNull.Value)
                                    workSheet.Cells[f, c].Value = r["agora"].ToString();

                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                c++;

                                //////if (r["tx_apell_cli"] != System.DBNull.Value)
                                //////    workSheet.Cells[f, c].Value = r["tx_apell_cli"].ToString();
                                //////c++;

                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = apellido;
                                }
                                else
                                {
                                    if (r["tx_apell_cli"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["tx_apell_cli"].ToString();
                                }
                                c++;

                                //DIAS_ESTADO_SAP
                                workSheet.Cells[f, c].Value = GetDiasEstado(r["cups20"].ToString());
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                c++;

                                //DIAS_ESTADO_KEE
                                workSheet.Cells[f, c].Value = "";c++;

                                //DIAS_ESTADO_GLOBAL
                                workSheet.Cells[f, c].Value = ""; c++;

                                //Recuperar Incidencia Facturacion
                                listaIncidenciasSAP = RelacionIncidenciasSAP.GetCups(r["cups20"].ToString() + aniomes);

                                if (listaIncidenciasSAP.Count == 0)
                                {
                                    IncidenciaFacturacion = "";
                                    Estado_FAC_SE = "";
                                    Titulo_FAC = "";
                                }
                                else
                                {
                                    foreach (var mm in listaIncidenciasSAP)
                                    {
                                        string[] cadena = mm.Split(';');
                                        IncidenciaFacturacion = cadena[0];
                                        Estado_FAC_SE = cadena[1];
                                        Titulo_FAC = cadena[2];
                                        break;
                                    }
                                }
                                //Fin recuperar incidencia facturacion


                                // Comprobamos con la tabla Relacion_INC_CUPS
                                ////// 1.Se relaciona el detalle del informe obtenido(campos cups20 y mes) con la tabla "Relacion_INC_CUPS"(campos cups y mes).Si para un cups se encuentra en esa relación una incidencia que afeca al mismo mes que está pendiente se deberá desechar el estado que tuviera el informe y actualizarlo por uno nuevo.
                                ////// 2.El estado a actualizar dependerá del campo "area" de la tabla "Relacion_INC_CUPS", pudiendo existir 3 opciones:
                                ////// --01.B09 Incidencia_Contratacion: se informará con este estado cuando encuentre para ese mes una incidencia en el area de contratacion, independientemente de que también haya otras para ese mismo mes en otras areas.
                                //////-- 01.B10 Incidencia_Medida: no estando en el caso anterior, se informará con este estado cuando encuentre para ese mes una incidencia en el area de medida, independientemente de que también haya otras para ese mismo mes en el area de facturacion.
                                ////// --01.B11 Incidencia_Facturacion: no estando en ninguno de los dos casos anteriores, se informará con este estado cuando encuentre para ese mes una incidencia en el area de Facturacion.
                                strSql = " select area, incidencia, estado_incidencia, fecha_apertura, prioridad_negocio, titulo, e_s_estado, Reincidente from Relacion_INC_CUPS"
                                         + " where cups='" + r["cups20"].ToString() + "' and Mes_pendiente='" + aniomes + "'"
                                         + " order by Fecha_apertura asc";

                                dbIncidencia = new MySQLDB(MySQLDB.Esquemas.MED);
                                command = new MySqlCommand(strSql, dbIncidencia.con);
                                inci = command.ExecuteReader();

                                areaincidencia = "";
                                incidencia = "";
                                estado_incidencia = "";
                                fecha_apertura = "";
                                prioridad_negocio = "";
                                titulo = "";
                                e_s_estado = "";
                                subestado_incidencia = "";
                                reincidente = "";
                                ExisteIncidencia = false;

                                while (inci.Read())
                                {
                                    ExisteIncidencia = true;
                                    if (areaincidencia == "")
                                    {
                                        areaincidencia = inci["area"].ToString();

                                        if (areaincidencia == "Contratacion")
                                            subestado_incidencia = "01.B09 Incidencia_Contratacion";
                                        if (areaincidencia == "Medida")
                                            subestado_incidencia = "01.B10 Incidencia_Medida";
                                        if (areaincidencia == "Facturacion")
                                            subestado_incidencia = "01.B11 Incidencia_Facturacion";

                                        incidencia = inci["incidencia"].ToString();
                                        estado_incidencia = inci["estado_incidencia"].ToString();
                                        fecha_apertura = inci["fecha_apertura"].ToString();
                                        prioridad_negocio = inci["prioridad_negocio"].ToString();
                                        titulo = inci["titulo"].ToString();
                                        e_s_estado = inci["e_s_estado"].ToString();
                                        reincidente = inci["Reincidente"].ToString();
                                    }
                                    else
                                    {
                                        if (areaincidencia != inci["area"].ToString())
                                        {
                                            areaincidencia = inci["area"].ToString();

                                            if (areaincidencia == "Contratacion")
                                                subestado_incidencia = subestado_incidencia + " - 01.B09 Incidencia_Contratacion";
                                            if (areaincidencia == "Medida")
                                                subestado_incidencia = subestado_incidencia + " - 01.B10 Incidencia_Medida";
                                            if (areaincidencia == "Facturacion")
                                                subestado_incidencia = subestado_incidencia + " - 01.B11 Incidencia_Facturacion";

                                            incidencia = incidencia + " - " + inci["incidencia"].ToString();
                                            estado_incidencia = estado_incidencia + " - " + inci["estado_incidencia"].ToString();
                                            fecha_apertura = fecha_apertura + " - " + inci["fecha_apertura"].ToString();
                                            prioridad_negocio = prioridad_negocio + " - " + inci["prioridad_negocio"].ToString();
                                            titulo = titulo + " - " + inci["titulo"].ToString();
                                            e_s_estado = e_s_estado + " - " + inci["e_s_estado"].ToString();
                                        }
                                    }
                                } // Fin  while (inci.Read())

                                dbIncidencia.CloseConnection();
                                if (ExisteIncidencia == true)
                                {
                                    //////c--;
                                    //////workSheet.Cells[f, c].Value = subestado_incidencia; c++;
                                    ///
                                    workSheet.Cells[f, c].Value = IncidenciaFacturacion;c++;
                                    workSheet.Cells[f, c].Value = Estado_FAC_SE; c++;
                                    workSheet.Cells[f, c].Value = Titulo_FAC; c++;
                                    workSheet.Cells[f, c].Value = incidencia; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = estado_incidencia; c++;
                                    workSheet.Cells[f, c].Value = fecha_apertura; c++;
                                    workSheet.Cells[f, c].Value = prioridad_negocio; c++;
                                    workSheet.Cells[f, c].Value = titulo; c++;
                                    workSheet.Cells[f, c].Value = e_s_estado; c++;
                                }
                                else
                                {
                                    workSheet.Cells[f, c].Value = IncidenciaFacturacion; c++;
                                    workSheet.Cells[f, c].Value = Estado_FAC_SE; c++;
                                    workSheet.Cells[f, c].Value = Titulo_FAC; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                }
                                ////////////////////////////////////////////////////////////////////////////////////////////////


                                //FH_ALTA_SALESFORCE
                                ////workSheet.Cells[f, c].Value = ""; c++;

                                if (r["fh_alta_crto"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_alta_crto"]).Date;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                else
                                {

                                }
                                c++;
                                // FIN FH_ALTA_SALESFORCE

                                //FECHA_ALTA_KEE*****************************************************
                                listaAltaKEE = pendienteWeb_B2B.GetCupsFechaAltaKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fec_act"]));

                                if (listaAltaKEE.Count == 0)
                                {
                                    workSheet.Cells[f, c].Value = ""; c++;
                                }
                                else
                                {
                                    if (listaAltaKEE[0] == "01/01/0001 0:00:00")
                                    {
                                        workSheet.Cells[f, c].Value = ""; c++;
                                    }
                                    else
                                    {

                                        foreach (var mm in listaAltaKEE)
                                        {
                                            workSheet.Cells[f, c].Value = mm; c++;
                                            break;
                                        }
                                        //////workSheet.Cells[f, c].Value = lista; c++;
                                    }
                                }

                                //FIN FH_ALTA_KEE*****************************************************

                                //////if (r["fh_alta_crto"] != System.DBNull.Value)
                                //////{
                                //////    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_alta_crto"]).Date;
                                //////    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                //////}
                                //////c++;


                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToDateTime(FechaAlta).Date;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                else
                                {
                                    if (r["fh_alta_crto"] != System.DBNull.Value)
                                    {
                                        workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_alta_crto"]).Date;
                                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    }
                                }
                                c++;


                                if (r["fh_baja_crto"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_baja_crto"]).Date;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                                }
                                else
                                {
                                    ////////if (r["fh_prev_fin_crto"] != System.DBNull.Value)
                                    ////////{
                                    ////////    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_prev_fin_crto"]).Date;
                                    ////////    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                                    ////////}
                                }
                                c++;

                                //Paco, buscamos la fecha de Baja de SAP ******************************************
                                CadenaAuxiliar = r["id_crto_ext"].ToString();
                                CadenaAuxiliar = CadenaAuxiliar.PadLeft(12, '0'); //Relleno con ceros a la izquierda hasta 12


                                fhBajaSap = FechaBajaSap.GetCupsFechaBajaSAP(CadenaAuxiliar);

                                if (fhBajaSap.Count == 0)
                                {
                                    workSheet.Cells[f, c].Value = ""; c++;
                                }
                                else
                                {
                                    foreach (var mm in fhBajaSap)
                                    {
                                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                        workSheet.Cells[f, c].Value = mm; c++;   
                                        break;
                                    }
                                }


                                //////strSql = "SELECT  max(cd_sec_crto), fh_baja from ed_owner.t_ed_h_sap_crto_front "
                                //////+ " WHERE id_crto_ext = '" + CadenaAuxiliar + "'"
                                //////+ " group by  fh_baja"
                                //////+ " order by  max(cd_sec_crto) desc";

                                //////dbRS = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                                //////commandRS = new OdbcCommand(strSql, dbRS.con);
                                //////rRS = commandRS.ExecuteReader();
                                //////while (rRS.Read())
                                //////{
                                //////    if (rRS["fh_baja"] != System.DBNull.Value)
                                //////        workSheet.Cells[f, c].Value = Convert.ToDateTime(rRS["fh_baja"]).Date;
                                //////    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                //////    break;
                                //////}
                                //////dbRS.CloseConnection();
                                //////c++;

                                //Fin - Fh_baja de SAP ***************************************************


                                //////if (r["cd_empr"] != System.DBNull.Value)
                                //////    workSheet.Cells[f, c].Value = r["cd_empr"].ToString();
                                //////c++;


                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = empresa;
                                }
                                else
                                {
                                    if (r["cd_empr"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["cd_empr"].ToString();
                                }
                                c++;

                                if (r["cd_segmento_ptg_Con_compor"] != System.DBNull.Value)
                                    workSheet.Cells[f, c].Value = r["cd_segmento_ptg_Con_compor"].ToString();
                                c++;

                                //////if (r["de_tp_cli"] != System.DBNull.Value)
                                //////    workSheet.Cells[f, c].Value = r["de_tp_cli"].ToString();
                                //////c++;


                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = cliente;
                                }
                                else
                                {
                                    if (r["de_tp_cli"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["de_tp_cli"].ToString();
                                }
                                c++;

                                //////if (r["cd_nif_cif_cli"] != System.DBNull.Value)
                                //////    workSheet.Cells[f, c].Value = r["cd_nif_cif_cli"].ToString();
                                //////c++;

                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = nif;
                                }
                                else
                                {
                                    if (r["cd_nif_cif_cli"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["cd_nif_cif_cli"].ToString();
                                }
                                c++;

                                //////if (r["fh_inicio_vers_crto"] != System.DBNull.Value)
                                //////{
                                //////    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_inicio_vers_crto"]).Date;
                                //////    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                //////}
                                //////c++;


                                if (Existetedhps == false)
                                {     
                                    workSheet.Cells[f, c].Value = Convert.ToDateTime(FechaInicio).Date;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                else
                                {
                                    if (r["fh_inicio_vers_crto"] != System.DBNull.Value)
                                    {
                                        workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_inicio_vers_crto"]).Date;
                                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    }
                                }
                                c++;

                                if (r["id_instalacion"] != System.DBNull.Value)
                                    workSheet.Cells[f, c].Value = r["id_instalacion"].ToString();
                                c++;

                                //////if (r["cd_tarifa_c"] != System.DBNull.Value)
                                //////    workSheet.Cells[f, c].Value = r["cd_tarifa_c"].ToString();
                                //////c++;


                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = Tarifa;
                                }
                                else
                                {
                                    if (r["cd_tarifa_c"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["cd_tarifa_c"].ToString();
                                }
                                c++;

                                //////if (r["cd_crto_comercial"] != System.DBNull.Value)
                                //////    workSheet.Cells[f, c].Value = r["cd_crto_comercial"].ToString();
                                //////c++;


                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = contrato;
                                }
                                else
                                {
                                    if (r["cd_crto_comercial"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["cd_crto_comercial"].ToString();
                                }
                                c++;

                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = Distribuidora;
                                }
                                else
                                {
                                    if (r["de_empr_distdora_nombre"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["de_empr_distdora_nombre"].ToString();
                                }
                                c++;

                                if (r["de_estado"] != System.DBNull.Value)
                                    workSheet.Cells[f, c].Value = r["de_estado"].ToString();
                                c++;

                                if (r["de_subestado"] != System.DBNull.Value)
                                    workSheet.Cells[f, c].Value = r["de_subestado"].ToString();
                                c++;

                                //if (r["lg_multimedida"] != System.DBNull.Value)
                                //{
                                //    workSheet.Cells[f, c].Value = r["lg_multimedida"].ToString();

                                //}
                                //else
                                //{
                                //    workSheet.Cells[f, c].Value = "N";
                                //}


                                //PRIUEBA
                                //////workSheet.Cells[f, c].Value = GetDiasEstado(r["cups20"].ToString());

                                //////workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //////c++;

                                if (r["TAM"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                }
                                c++;



                                // Paco - añadimos la ultima fecha desde y hasta facturada t_ed_h_sap_facts
                                //workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                                //////if (pendienteSAPKEE.GetCupsInicioUltimaFacturada(r["cups20"].ToString()) is null)
                                //////{
                                //////    workSheet.Cells[f, c].Value = "";
                                //////}
                                //////else
                                //////{
                                //////    foreach (var mm in pendienteSAPKEE.GetCupsInicioUltimaFacturada(r["cups20"].ToString()))
                                //////    {
                                //////        workSheet.Cells[f, c].Value = mm;
                                //////    }
                                //////    //////workSheet.Cells[f, c].Value = pendienteSAPKEE.GetCupsInicioUltimaFacturada(r["cups20"].ToString());
                                //////}


                                // Paco - añadimos la ultima fecha desde y hasta facturada t_ed_h_sap_facts
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                                List<string> listaInicio = new List<string>();
                                listaInicio = pendienteSAPKEE.GetCupsInicioUltimaFacturada(r["cups20"].ToString());
                                if (listaInicio != null)
                                {
                                    if (listaInicio.Count == 0)
                                    {
                                        workSheet.Cells[f, c].Value = ""; c++;
                                    }
                                    else
                                    {
                                        if (listaInicio[0] == "01/01/0001 0:00:00")
                                        {
                                            workSheet.Cells[f, c].Value = ""; c++;
                                        }
                                        else
                                        {

                                            foreach (var mm in listaInicio)
                                            {
                                                workSheet.Cells[f, c].Value = mm; c++;
                                                break;
                                            }
                                            ////workSheet.Cells[f, c].Value = lista; c++;
                                        }
                                    }
                                }
                                else
                                {
                                    workSheet.Cells[f, c].Value = ""; c++;
                                }

                                //////if (workSheet.Cells[f, c].Value == null)
                                //////{
                                //////    workSheet.Cells[f, c].Value = "";
                                //////}
                                //////c++;
                                //workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                                // Paco - añadimos la ultima fecha desde y hasta facturada t_ed_h_sap_facts
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                                List<string> listaFin = new List<string>();
                                listaFin = pendienteSAPKEE.GetCupsFinalUltimaFacturada(r["cups20"].ToString());
                                if (listaFin != null)
                                {
                                    if (listaFin.Count == 0)
                                    {
                                        workSheet.Cells[f, c].Value = ""; c++;
                                    }
                                    else
                                    {
                                        if (listaFin[0] == "01/01/0001 0:00:00")
                                        {
                                            workSheet.Cells[f, c].Value = ""; c++;
                                        }
                                        else
                                        {

                                            foreach (var mm in listaFin)
                                            {
                                                workSheet.Cells[f, c].Value = mm; c++;
                                                break;
                                            }
                                            ////workSheet.Cells[f, c].Value = lista; c++;
                                        }
                                    }
                                }
                                else
                                {
                                    workSheet.Cells[f, c].Value = ""; c++;
                                }

                                subestadosKronos.existe = false;
                                if (subestados_sap.AreaResponsableMedida(r["cd_subestado"].ToString()))
                                {

                                    subestadosKronos.existe = false;

                                    //////if (r["cups20"].ToString()== "PT0002000111610262EM" )
                                    //////{
                                    //////    MessageBox.Show("");
                                    //////}

                                    //////if (r["cups20"].ToString() == "PT0002000116086035HR")
                                    //////{
                                    //////    MessageBox.Show("");
                                    //////}


                                    //IMPORTANTE!!!! Limpiamos el area para que no queda una antigua en la revisión de incidencias
                                    subestadosKronos.area_responsable  = null;
                                    
                                    if (multipunto == "N")
                                    {
                                        subestadosKronos.GetEstadoKEEDetalle(pendienteWeb_B2B.GetCupsDetalle(r["cups20"].ToString(),
                                            Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]), Convert.ToDateTime(r["fec_act"])));
                                    }
                                    else
                                    {
                                        subestadosKronos.GetEstadoKEEDetalleMultipunto(pendienteWeb_B2B.GetCupsDetalle(r["cups20"].ToString(),
                                            Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]), Convert.ToDateTime(r["fec_act"])));
                                    }


                                    if (subestadosKronos.existe)
                                    {
                                        BuscaCadena = subestadosKronos.descripcion_estado;
                                        List<string> lista = new List<string>();

                                        if (BuscaCadena.IndexOf("Discrepancia") >= 0)
                                        {
                                            if (BuscaCadena.IndexOf("Periodo SAP encontrado parcialmente en el informe pendiente de KEE") >= 0)
                                            {
                                                workSheet.Cells[f, c].Value = subestadosKronos.estado_periodo; c++;
                                                workSheet.Cells[f, c].Value = subestadosKronos.area_responsable; c++;
                                                workSheet.Cells[f, c].Value = "01.B05 Error Sistemas KEE - SAP - Incoherencia periodos KEE-SAP"; c++;
                                                // Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el estado de KEE lo he guardado en descripcion_subestado
                                                //en lugar de en descripcion_estado
                                                workSheet.Cells[f, c].Value = subestadosKronos.descripcion_subestado; c++;
                                                string[] cadena = subestadosKronos.descripcion_estado.Split(';');
                                                string[] cadenaAux = cadena[2].Split('/');
                                                workSheet.Cells[f, c].Value = cadenaAux[0].Substring(1, 8).Substring(6, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(4, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(0, 4); c++;
                                                workSheet.Cells[f, c].Value = cadenaAux[1].Substring(0, 8).Substring(6, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(4, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(0, 4); c++;
                                            }
                                            //|| BuscaCadena.IndexOf("Periodos del pendiente de KEE contenidos en las fechas de SAP") >= 0 || BuscaCadena.IndexOf("Periodo SAP contenido dentro de un informe pendiente de KEE") >= 0
                                            else
                                            {
                                                if (BuscaCadena.IndexOf("Periodos del pendiente de KEE contenidos en las fechas de SAP") >= 0 || BuscaCadena.IndexOf("Periodo SAP contenido dentro de un informe pendiente de KEE") >= 0)
                                                {
                                                    workSheet.Cells[f, c].Value = subestadosKronos.estado_periodo; c++;
                                                    workSheet.Cells[f, c].Value = subestadosKronos.area_responsable; c++;
                                                    // Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el subestado de KEE lo he guardado en temporal de momento
                                                    workSheet.Cells[f, c].Value = subestadosKronos.temporal; c++;
                                                    // Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el estado de KEE lo he guardado en descripcion_subestado
                                                    //en lugar de en descripcion_estado
                                                    workSheet.Cells[f, c].Value = subestadosKronos.descripcion_subestado; c++;
                                                    string[] cadena = subestadosKronos.descripcion_estado.Split(';');
                                                    string[] cadenaAux = cadena[2].Split('/');
                                                    workSheet.Cells[f, c].Value = cadenaAux[0].Substring(1, 8).Substring(6, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(4, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(0, 4); c++;
                                                    workSheet.Cells[f, c].Value = cadenaAux[1].Substring(0, 8).Substring(6, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(4, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(0, 4); c++;

                                                }
                                                else
                                                {
                                                    //si no está el CUPS en la tabla "ed_owner.T_ED_H_GEST_DIAR_PS_B2B" he obtenido: "Discrepancia: No existe el cups en el informe del pendiente de KEE"
                                                    // el cups está en la tabla, pero dentro de los periodos que hay no no hay ningún periodo que coincida ni total ni parcialmente
                                                    //      Marcado: "Discrepancia: No existen fechas en el informe pendiente de KEE para el periodo;(" + Convert.ToString(fecha_desde.ToString("yyyyMMdd")) + "/" + Convert.ToString(fecha_hasta.ToString("yyyyMMdd")) + ")")

                                                    if (BuscaCadena.IndexOf("No existe el cups en el informe del pendiente de KEE") >= 0)
                                                    {
                                                        workSheet.Cells[f, c].Value = ""; c++;
                                                        workSheet.Cells[f, c].Value = ""; c++;
                                                        workSheet.Cells[f, c].Value = "01.B07 Error Sistemas KEE-SAP - Pdte SAP-No existe en KEE"; c++;
                                                        workSheet.Cells[f, c].Value = ""; c++;
                                                        workSheet.Cells[f, c].Value = ""; c++;
                                                        workSheet.Cells[f, c].Value = ""; c++;

                                                    }
                                                    else
                                                    {

                                                        if (BuscaCadena.IndexOf("No existen fechas en el informe pendiente de KEE para el periodo") >= 0)
                                                        {
                                                            workSheet.Cells[f, c].Value = ""; c++;
                                                            workSheet.Cells[f, c].Value = ""; c++;
                                                            workSheet.Cells[f, c].Value = "01.B06 Error Sistemas KEE-SAP - Pdte SAP-No pdte KEE"; c++;
                                                            workSheet.Cells[f, c].Value = ""; c++;
                                                            workSheet.Cells[f, c].Value = ""; c++;
                                                            workSheet.Cells[f, c].Value = ""; c++;

                                                        }
                                                        else
                                                        {
                                                            workSheet.Cells[f, c].Value = ""; c++;
                                                            workSheet.Cells[f, c].Value = ""; c++;
                                                            workSheet.Cells[f, c].Value = ""; c++;
                                                            workSheet.Cells[f, c].Value = ""; c++;
                                                            workSheet.Cells[f, c].Value = ""; c++;
                                                            workSheet.Cells[f, c].Value = ""; c++;
                                                        }

                                                    } //Fin  if (BuscaCadena.IndexOf("No existe el cups en el informe del pendiente de KEE") >= 0)

                                                } // fin  if (BuscaCadena.IndexOf("Periodos del pendiente de KEE contenidos en las fechas de SAP") >= 0 || BuscaCadena.IndexOf("Periodo SAP contenido dentro de un informe pendiente de KEE") >= 0)

                                            } // Fin  if (BuscaCadena.IndexOf("Periodo SAP encontrado parcialmente en el informe pendiente de KEE") >= 0 )

                                            //////lista = pendienteWeb_B2B.GetCupsFinKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]));
                                            lista = pendienteWeb_B2B.GetCupsFinKEENew(r["cups20"].ToString(), Convert.ToDateTime(r["fec_act"]));

                                            if (lista.Count == 0)
                                            {
                                                workSheet.Cells[f, c].Value = ""; c++;
                                            }
                                            else
                                            {
                                                if (lista[0] == "01/01/0001 0:00:00")
                                                {
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                }
                                                else
                                                {

                                                    foreach (var mm in lista)
                                                    {
                                                        workSheet.Cells[f, c].Value = mm; c++;
                                                        break;
                                                    }
                                                    //////workSheet.Cells[f, c].Value = lista; c++;
                                                }
                                            }
                                            workSheet.Cells[f, c].Value = subestadosKronos.descripcion_estado; c++;
                                            //Subestado Global
                                            if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                            {
                                                workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoKEE].Value; c++;
                                            }
                                            else
                                            {
                                                workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoSAP].Value; c++;
                                            }

                                            //
                                        }
                                        else
                                        {
                                            workSheet.Cells[f, c].Value = subestadosKronos.estado_periodo; c++;
                                            workSheet.Cells[f, c].Value = subestadosKronos.area_responsable; c++;
                                            workSheet.Cells[f, c].Value = subestadosKronos.descripcion_subestado; c++;
                                            workSheet.Cells[f, c].Value = subestadosKronos.descripcion_estado; c++;
                                            // fh_desde_kee y fh_hasta_kee coincide con fecha desde y hasta de sap, equivalencia de fechas
                                            if (r["fh_desde"] != System.DBNull.Value)
                                            {
                                                workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_desde"]).Date;
                                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                            }
                                            c++;

                                            if (r["fh_hasta"] != System.DBNull.Value)
                                            {
                                                workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_hasta"]).Date;
                                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                            }
                                            c++;
                                            //////lista = pendienteWeb_B2B.GetCupsFinKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]));
                                            lista = pendienteWeb_B2B.GetCupsFinKEENew(r["cups20"].ToString(), Convert.ToDateTime(r["fec_act"]));
                                            if (lista.Count == 0)
                                            {
                                                workSheet.Cells[f, c].Value = ""; c++;
                                            }
                                            else
                                            {
                                                if (lista[0] == "01/01/0001 0:00:00")
                                                {
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                }
                                                else
                                                {

                                                    foreach (var mm in lista)
                                                    {
                                                        workSheet.Cells[f, c].Value = mm; c++;
                                                        break;
                                                    }
                                                    //////workSheet.Cells[f, c].Value = lista; c++;
                                                }
                                            }
                   
                                            workSheet.Cells[f, c].Value = ""; c++;

                                            //Subestado Global
                                            if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                            {
                                                workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoKEE].Value; c++;
                                            }
                                            else
                                            {
                                                workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoSAP].Value; c++;
                                            }
                                            //Fin Subestado Global
                                        }

                                        //Paco 21/05/2024 Calculo del numero dias KEE en Kronos ********************
                                        if (blnLanzarDiasKEE == true)
                                        {
                                            //////if (BuscaCadena.IndexOf("Discrepancia") >= 0)
                                            //////{
                                            //////    DiasExistenciaKEE = FuncionDiasKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), subestadosKronos.descripcion_subestado);
                                            //////}
                                            //////else
                                            //////{
                                            //////    DiasExistenciaKEE = FuncionDiasKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), subestadosKronos.descripcion_estado);
                                            //////}
                                            ///
                                            DiasExistenciaKEE = FuncionDiasKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), subestadosKronos.estado_periodo);
                                        }

                                        else
                                        {
                                            /// Para que recupere los dias en ejecuciones sucesivas el mismo día.
                                            //////if (BuscaCadena.IndexOf("Discrepancia") >= 0)
                                            //////{
                                            //////    DiasExistenciaKEE = FuncionDiasKEERecuperarDiasTabla(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), subestadosKronos.descripcion_subestado);
                                            //////}
                                            //////else
                                            //////{

                                            //////    DiasExistenciaKEE = FuncionDiasKEERecuperarDiasTabla(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), subestadosKronos.descripcion_estado);

                                            //////}
                                            ///
                                            DiasExistenciaKEE = FuncionDiasKEERecuperarDiasTabla(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), subestadosKronos.estado_periodo);

                                        }
                                        //Fin Calculo del numero dias KEE en Kronos  ***************


                                    }
                                    else
                                    {
                                        workSheet.Cells[f, c].Value = ""; c++;
                                        workSheet.Cells[f, c].Value = ""; c++;
                                        workSheet.Cells[f, c].Value = ""; c++;
                                        workSheet.Cells[f, c].Value = ""; c++;
                                        workSheet.Cells[f, c].Value = ""; c++;
                                        workSheet.Cells[f, c].Value = ""; c++;
                                        workSheet.Cells[f, c].Value = ""; c++;
                                        workSheet.Cells[f, c].Value = ""; c++;
                                        //Subestado Global
                                        if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                        {
                                            workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoKEE].Value; c++;
                                        }
                                        else
                                        {
                                            workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoSAP].Value; c++;
                                        }
                                        //Fin Subestado Global
                                    } //Fin if (subestadosKronos.existe)
                                }


                                else
                                {
                                    //System.Diagnostics.Debug.WriteLine(r["cups20"].ToString() + "-" + r["de_subestado"].ToString() + "-" + subestadosKronos.descripcion_subestado + "-" + subestadosKronos.descripcion_estado);
                                    //Para que llegue hasta la última columna aunque no exista en Kronos
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    // Fecha fin KEE (se pinta aunque no sea uno de los estados marcados como KEE
                                    // workSheet.Cells[f, c].Value = ""; c++;
                                    List<string> listaAux = new List<string>();
                                    listaAux = pendienteWeb_B2B.GetCupsFinKEENew(r["cups20"].ToString(), Convert.ToDateTime(r["fec_act"]));
                                    if (listaAux.Count == 0)
                                    {
                                        workSheet.Cells[f, c].Value = ""; c++;
                                    }
                                    else
                                    {
                                        if (listaAux[0] == "01/01/0001 0:00:00")
                                        {
                                            workSheet.Cells[f, c].Value = ""; c++;
                                        }
                                        else
                                        {

                                            foreach (var mm in listaAux)
                                            {
                                                workSheet.Cells[f, c].Value = mm; c++;
                                                break;
                                            }
                                            //////workSheet.Cells[f, c].Value = lista; c++;
                                        }
                                    }
                                    workSheet.Cells[f, c].Value = ""; c++;


                                    //Subestado Global
                                    if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                    {
                                        workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoKEE].Value; c++;
                                    }
                                    else
                                    {
                                        workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoSAP].Value; c++;
                                    }

                                } // Fin  if (subestados_sap.AreaResponsableMedida(r["cd_subestado"].ToString()))

                                //Calculo DIAS ESTADO GLOBAL *************************
                                if (blnLanzarDiasKEE == true) // Sólo se actualizan los datos en la tabla la primera vez, por si hay ejecuciones sucesivas
                                {
                                    if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                    {
                                        DiasExistenciaKEEGlobal = FuncionDiasKEEGlobal(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), workSheet.Cells[f, PosicionSubEstadoKEE].Value.ToString());
                                    }
                                    else
                                    {
                                        DiasExistenciaKEEGlobal = FuncionDiasKEEGlobal(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), workSheet.Cells[f, PosicionSubEstadoSAP].Value.ToString());
                                    }
                                }
                                else
                                {
                                    if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                    {
                                        DiasExistenciaKEEGlobal = FuncionDiasKEERecuperarDiasTablaGlobal(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), workSheet.Cells[f, PosicionSubEstadoKEE].Value.ToString());
                                    }
                                    else
                                    {
                                        DiasExistenciaKEEGlobal = FuncionDiasKEERecuperarDiasTablaGlobal(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), workSheet.Cells[f, PosicionSubEstadoSAP].Value.ToString());
                                    }
                                }
                                //FIN DIAS ESTADO GLOBAL*******************************


                                if (areaincidencia != "" & subestadosKronos.existe == true)
                                {
                                    Auxiliar = Reincidente(areaincidencia, subestadosKronos.area_responsable);
                                    if (Auxiliar != "")
                                    {
                                        ////c--;
                                        workSheet.Cells[f, PosicionSubEstadoGlobal].Value = Auxiliar;
                                        workSheet.Cells[f, PosicionReincidente].Value = reincidente;
                                    }
                                }

                                //Fin Paco 09/05/2024***********************************************************************************************


                                Auxiliar = "";
                                Auxiliar = EstadosAreas(workSheet.Cells[f, PosicionSubEstadoGlobal].Value.ToString(), cups20, aniomes.ToString(), 1,0, false);

                                if (Auxiliar != "")
                                {
                                    string[] CadenaEstados = Auxiliar.Split(';');
                                    workSheet.Cells[f, c].Value = CadenaEstados[0]; c++;
                                    workSheet.Cells[f, c].Value = CadenaEstados[1]; c++;
                                    workSheet.Cells[f, c].Value = CadenaEstados[2]; c++;
                                }
                                else
                                {
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;

                                    if (SubestadosNoEncontrados == "")
                                    {
                                        SubestadosNoEncontrados = "Hoja MT-BTE: " + cups20 + " - " + Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd") + " - " + Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd") + " - " + r["de_estado"].ToString() + " -" + r["de_subestado"].ToString() + "- No existe el subestado en la tabla paramétrica";

                                    }
                                    else
                                    {
                                        SubestadosNoEncontrados = SubestadosNoEncontrados + ";Hoja MT-BTE: " + cups20 + " - " + Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd") + " - " + Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd") + " - " + r["de_estado"].ToString() + " -" + r["de_subestado"].ToString() + "- No existe el subestado en la tabla paramétrica";
                                    }
                                }

                                //// Tratamiento multipuntos
                                workSheet.Cells[f, c].Value = multipunto; c++;

                                // "Nº DÍAS ESTADO KEE"
                                if (subestadosKronos.existe == true)
                                {
                                    workSheet.Cells[f, PosicionDiasEstadoKEE].Value = DiasExistenciaKEE; c++;
                                }

                                // Pego Nº DIAS ESTADO GLOBAL, va a continuación de DIAS ESTADO KEE
                                workSheet.Cells[f, PosicionDiasEstadoKEE + 1].Value = DiasExistenciaKEEGlobal; c++;

                            } //Fin if (Segmentos.IndexOf(r["cd_segmento_ptg"].ToString()) > 0)
                        } // fin  if (r["cd_segmento_ptg"] != System.DBNull.Value)

                    } // Fin  while (r.Read())
                    db.CloseConnection();

                    intConteoMTBTEDetalle = f-1;

                    if (blnLanzarDiasKEE == true) {
                        ///todos los que están activos y  no se ha marcado PROCESADO_DIARIO = 'S', son puntos que se cierran, actualizo la FH_CIERRE
                        strSql = "update kee_Dias  set"
                                   + " FH_CIERRE = '" + DateTime.Now.ToString("yyyy-MM-dd") + "',"
                                   + " FH_ACTUALIZACION = '" + DateTime.Now.ToString("yyyy-MM-dd") + "',"
                                   + " ACTIVO= False"
                                   + " WHERE  Activo = true"
                                   + " AND FH_ACTUALIZACION < '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";

                        db = new MySQLDB(MySQLDB.Esquemas.MED);
                        command = new MySqlCommand(strSql, db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();

                    }

                    ////////13/02/2024 Paco, mover columna de fecha baja de Kronos despues de fecha de baja de SAP
                    //////workSheet.InsertColumn(PosicionBajaSap, 1); //Inserta en la posicion 12 una columna en blanco(detrás de FH_BAJA_SAP)
                    //////workSheet.Cells[1, PosicionBajaKEE, f, PosicionBajaKEE].Copy(workSheet.Cells[1, PosicionBajaSap, f, PosicionBajaSap]); //Copia la columna 31 (Fecha BAJA KEE) en la posición creada anteriormente
                    //////workSheet.DeleteColumn(PosicionBajaKEE); //Borra la columna 31 Fecha BAJA KEE

                    //////allCells = workSheet.Cells[1, 1, f, c];
                    //////allCells.AutoFitColumns();

                    headerCells = workSheet.Cells[1, 1, 1, c];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;

                    //08/02/2024 Paco, mover columna de fecha baja de Kronos despues de fecha de baja de SAP
                    workSheet.InsertColumn(PosicionBajaSap, 1); //Inserta en la posicion 12 una columna en blanco(detrás de FH_BAJA_SAP)
                    workSheet.Cells[1, PosicionBajaKEE, f, PosicionBajaKEE].Copy(workSheet.Cells[1, PosicionBajaSap, f, PosicionBajaSap]); //Copia la columna 31 (Fecha BAJA KEE) en la posición creada anteriormente
                    workSheet.DeleteColumn(PosicionBajaKEE); //Borra la columna 31 Fecha BAJA KEE

                    //Pegamos Subestadoglobal, estado_global y estado global a reportar despues de FH_HASTA *** (!!!OJO, he metido ya por medio FECHA BAJA KEE (Sumo una más)!!)****
                    workSheet.InsertColumn(PosicionPegarSubEstadoGlobal, 1); //Inserta en la posicion 5 una columna en blanco(detrás de FH_HASTA)
                    workSheet.Cells[1, PosicionSubEstadoGlobal + 1, f, PosicionSubEstadoGlobal + 1].Copy(workSheet.Cells[1, PosicionPegarSubEstadoGlobal, f, PosicionPegarSubEstadoGlobal]);

                    workSheet.InsertColumn(PosicionPegarSubEstadoGlobal + 1, 1); //inserto en posición 6 (ya he pegado subestadoglobal)
                    workSheet.Cells[1, PosicionSubEstadoGlobal + 3, f, PosicionSubEstadoGlobal + 3].Copy(workSheet.Cells[1, PosicionPegarSubEstadoGlobal + 1, f, PosicionPegarSubEstadoGlobal + 1]);

                    workSheet.InsertColumn(PosicionPegarSubEstadoGlobal + 2, 1); //inserto en posición 7 (ya he pegado subestadoglobal)
                    workSheet.Cells[1, PosicionSubEstadoGlobal + 5, f, PosicionSubEstadoGlobal + 5].Copy(workSheet.Cells[1, PosicionPegarSubEstadoGlobal + 2, f, PosicionPegarSubEstadoGlobal + 2]);

                    // borro las originales (tres a la derecha, he insertado tres columnas)
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal + 3);
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal + 3);
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal + 3);


                    // He quitado  lo que hay entre discrepancias y multipuntos, las posiciones me valen, ahora he metido los globales por delante
                    workSheet.InsertColumn(PosicionBajaKEE + 3, 1);
                    workSheet.Cells[1, PosicionSubEstadoGlobal + 4, f, PosicionSubEstadoGlobal + 4].Copy(workSheet.Cells[1, PosicionBajaKEE + 3, f, PosicionBajaKEE + 3]); //Copia la columna 31 (Fecha BAJA KEE) en la posición creada anteriormente
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal + 4);

                    //La columna AREA la pongo detrás de fh_hasta
                    workSheet.InsertColumn(PosicionPegarSubEstadoGlobal, 1);
                    // Ya he quitado lo que originalmente estaba detras de posiciónSubEstado Global, sólo queda el area en esa poscion
                    workSheet.Cells[1, PosicionSubEstadoGlobal + 3, f, PosicionSubEstadoGlobal + 3].Copy(workSheet.Cells[1, PosicionPegarSubEstadoGlobal, f, PosicionPegarSubEstadoGlobal]);
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal + 3);

                    //El final es discrepancias seguido de multipuntos, los tengo que cambiar de orden --> discrepancia=48 y multipunto=49 después de las inserciones de columna
                    //no lo puedo hacer durante el proceso y poner multipuntos antes, ya que las discrepancias es un proceso aparte y separado
                    //Busco donde ha quedado discrepancias
                    intColumna = 1;
                    while (workSheet.Cells[1, intColumna].Value != "" & workSheet.Cells[1, intColumna].Value != "Discrepancias")
                    {
                        intColumna = intColumna + 1;
                    }
                    workSheet.InsertColumn(intColumna, 1); //Inserta  columna en blanco delante de discrepancias
                    workSheet.Cells[1, intColumna + 2, f, intColumna + 2].Copy(workSheet.Cells[1, intColumna, f, intColumna]); //copio multipunto en la columna que he insertado
                    workSheet.DeleteColumn(intColumna + 2); //Borra la columna original de multipunto


                    //workSheet.Cells[1,5,100,5].Copy(workSheet.Cells[1,2,100,2]);
                    //Copia la columna 5 en la columna 2 Básicamente Source.Copy(Destino) Esto solo copiaría las primeras 100 filas.

                    allCells = workSheet.Cells[1, 1, f, c + 7]; //Pongo +7 para las columnas de incidencias
                    allCells.AutoFitColumns();

                    #endregion

                    #region Detalle POR BTN

                    f = 1;
                    c = 1;

                    workSheet = excelPackage.Workbook.Worksheets.Add("Detalle POR BTN");

                    headerCells = workSheet.Cells[1, 1, 1, 32];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;
                    allCells = workSheet.Cells[1, 1, 50, 50];


                    //Ponemos columnas directamente
                    workSheet.Cells[f, c].Value = "CUPS20"; c++;
                    workSheet.Cells[f, c].Value = "PERIODO"; c++;
                    workSheet.Cells[f, c].Value = "MES"; c++;
                    workSheet.Cells[f, c].Value = "FH_DESDE"; c++;
                    workSheet.Cells[f, c].Value = "FH_HASTA"; c++;
                    PosicionPegarSubEstadoGlobal = c;
                    workSheet.Cells[f, c].Value = "IMPORTE PDTE FACTURAR"; c++;
                    workSheet.Cells[f, c].Value = "MESES PDTES FACTURAR"; c++;
                    workSheet.Cells[f, c].Value = "ÁGORA"; c++;
                    workSheet.Cells[f, c].Value = "CLIENTE"; c++;
                    workSheet.Cells[f, c].Value = "DIAS_ESTADO_SAP"; c++;
                    PosicionDiasEstadoKEE = c;
                    workSheet.Cells[f, c].Value = "DIAS_ESTADO_KEE"; c++;
                    workSheet.Cells[f, c].Value = "DIAS_ESTADO_GLOBAL"; c++;
                    workSheet.Cells[f, c].Value = "Incidencia Facturacion"; c++;
                    workSheet.Cells[f, c].Value = "Estado_Fac_SE"; c++;
                    workSheet.Cells[f, c].Value = "Titulo_Fac"; c++;
                    workSheet.Cells[f, c].Value = "Incidencia Medida"; c++;
                    PosicionReincidente = c;
                    workSheet.Cells[f, c].Value = "Reincidente"; c++;
                    workSheet.Cells[f, c].Value = "Estado incidencia"; c++;
                    workSheet.Cells[f, c].Value = "Fecha_Apertura"; c++;
                    workSheet.Cells[f, c].Value = "Prioridad"; c++;
                    workSheet.Cells[f, c].Value = "Titulo"; c++;
                    workSheet.Cells[f, c].Value = "E_S_Estado"; c++;
                    workSheet.Cells[f, c].Value = "FH_ALTA_SALESFORCE"; c++;
                    workSheet.Cells[f, c].Value = "FH_ALTA_KEE"; c++;
                    workSheet.Cells[f, c].Value = "FH_ALTA_SAP"; c++;
                    workSheet.Cells[f, c].Value = "FH_BAJA_SALESFORCE"; c++;
                    PosicionBajaSap = c;
                    workSheet.Cells[f, c].Value = "FH_BAJA_SAP"; c++;
                    workSheet.Cells[f, c].Value = "EMPRESA"; c++;
                    workSheet.Cells[f, c].Value = "SEGMENTO"; c++;
                    workSheet.Cells[f, c].Value = "TIPO CLIENTE"; c++;
                    workSheet.Cells[f, c].Value = "NIF"; c++;
                    workSheet.Cells[f, c].Value = "FPSERCON"; c++;
                    workSheet.Cells[f, c].Value = "Nº INSTALACIÓN"; c++;
                    workSheet.Cells[f, c].Value = "TARIFA"; c++;
                    workSheet.Cells[f, c].Value = "CONTRATO"; c++;
                    workSheet.Cells[f, c].Value = "DISTRIBUIDORA"; c++;
                    workSheet.Cells[f, c].Value = "ESTADO"; c++;
                    PosicionSubEstadoSAP = c;
                    workSheet.Cells[f, c].Value = "SUBESTADO"; c++;
                    //////workSheet.Cells[f, c].Value = "DIAS_ESTADO"; c++;
                    workSheet.Cells[f, c].Value = "TAM"; c++;
                    workSheet.Cells[f, c].Value = "ULT FH DESDE FACTURADA"; c++;
                    workSheet.Cells[f, c].Value = "ULT FH HASTA FACTURADA"; c++;
                    workSheet.Cells[f, c].Value = "Estado periodo KEE"; c++;
                    workSheet.Cells[f, c].Value = "Área responsable KEE"; c++;
                    PosicionSubEstadoKEE = c;
                    workSheet.Cells[f, c].Value = "Subestado KEE"; c++;
                    workSheet.Cells[f, c].Value = "Estado KEE"; c++;
                    workSheet.Cells[f, c].Value = "FH_DESDE_KEE"; c++;
                    workSheet.Cells[f, c].Value = "FH_HASTA_KEE"; c++;
                    workSheet.Cells[f, c].Value = "Fecha BAJA KEE"; c++;
                    PosicionBajaKEE = c;
                    workSheet.Cells[f, c].Value = "Discrepancias"; c++;
                    PosicionSubEstadoGlobal = c;
                    workSheet.Cells[f, c].Value = "Subestado global"; c++;
                    workSheet.Cells[f, c].Value = "ESTADO GLOBAL"; c++;
                    workSheet.Cells[f, c].Value = "ESTADO GLOBAL A REPORTAR"; c++;
                    workSheet.Cells[f, c].Value = "AREA ESTADO"; c++;

                    //////workSheet.Cells[f, c].Value = "Nº DÍAS ESTADO KEE"; c++;
                    workSheet.Cells[f, c].Value = "Nº DÍAS PTE KEE"; c++;

                    strSql = DetalleExcel(fecha_informe, lista_empresas_PT, lista_segmentos_BTN);

                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();


                    Segmentos = "BTN";

                    while (r.Read())
                    {
                        ///07-11-2024
                        if (r["cd_segmento_ptg_Con_compor"] != System.DBNull.Value)
                        {
                            if (r["cd_segmento_ptg_Con_compor"].ToString() == "BTN")
                            {
                        //Fin 07/11/2024

                                f++;
                                c = 1;

                                empresa = "";
                                nif = "";
                                cliente = "";
                                apellido = "";
                                FechaAlta = null;
                                FechaInicio = null;
                                Tarifa = "";
                                contrato = "";
                                Distribuidora = "";
                                cups20 = "";
                                NumeroDiasKEE = 0;
                                meses_pdtes = 0;

                                if (r["cups20"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = r["cups20"].ToString();
                                    cups20= r["cups20"].ToString();
                                }
                                c++;

                                ///
                                workSheet.Cells[f, c].Value = "";c++;

                                if (r["mes"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["mes"]);
                                    aniomes = Convert.ToInt32(r["mes"]);
                                    fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                        Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                                }
                                c++;

                                if (r["fh_desde"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_desde"]).Date;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (r["fh_hasta"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_hasta"]).Date;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (r["mes"] != System.DBNull.Value && r["TAM"] != System.DBNull.Value)
                                {
                                    //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                                    meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["tam"]) * meses_pdtes;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                                }
                                else
                                {
                                    c++;
                                }

                                if (r["mes"] != System.DBNull.Value)
                                {
                                    //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                                    meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                                    workSheet.Cells[f, c].Value = meses_pdtes;
                                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                }
                                c++;

                                if (meses_pdtes > 1)
                                {
                                    workSheet.Cells[f, 2].Value = ">1P";
                                }
                                else
                                {
                                    workSheet.Cells[f, 2].Value = "1P";
                                }

                                if (r["agora"] != System.DBNull.Value)
                                    workSheet.Cells[f, c].Value = r["agora"].ToString();

                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                c++;


                                //Para controlar los blancos, miramos en t_ed_h_ps_hist
                                if (r["cd_empr"] != System.DBNull.Value)
                                {
                                    Existetedhps = true;
                                }
                                else
                                {
                                    Existetedhps = false;

                                    strSql = "SELECT cd_empr,  cd_nif_cif_cli, de_tp_cli, tx_apell_cli, fh_alta_crto, fh_inicio_vers_crto"
                                    + " ,ps.cd_tarifa_c, ps.cd_crto_comercial, ps.de_empr_distdora_nombre"
                                    + " FROM cont.t_ed_h_ps_pt_hist ps"
                                    + " WHERE cups20 = '" + r["cups20"].ToString() + "'"
                                    + " AND created_date = (SELECT MAX(created_date) FROM cont.t_ed_h_ps_pt_hist"
                                    + " WHERE cups20 = '" + r["cups20"].ToString() + "')";

                                    dbAux = new MySQLDB(MySQLDB.Esquemas.GBL);
                                    commandAux = new MySqlCommand(strSql, dbAux.con);
                                    rAux = commandAux.ExecuteReader();
                                    while (rAux.Read())
                                    {
                                        if (rAux["cd_empr"] != System.DBNull.Value)
                                            empresa = rAux["cd_empr"].ToString();
                                        if (rAux["de_tp_cli"] != System.DBNull.Value)
                                            cliente = rAux["de_tp_cli"].ToString();
                                        if (rAux["cd_nif_cif_cli"] != System.DBNull.Value)
                                            nif = rAux["cd_nif_cif_cli"].ToString();
                                        if (rAux["tx_apell_cli"] != System.DBNull.Value)
                                            apellido = rAux["tx_apell_cli"].ToString();
                                        if (rAux["fh_alta_crto"] != System.DBNull.Value)
                                            FechaAlta = Convert.ToDateTime(rAux["fh_alta_crto"]).Date;
                                        if (rAux["fh_inicio_vers_crto"] != System.DBNull.Value)
                                            FechaInicio = Convert.ToDateTime(rAux["fh_inicio_vers_crto"]).Date;
                                        if (rAux["cd_tarifa_c"] != System.DBNull.Value)
                                            Tarifa = rAux["cd_tarifa_c"].ToString();
                                        if (rAux["cd_crto_comercial"] != System.DBNull.Value)
                                            contrato = rAux["cd_crto_comercial"].ToString();
                                        if (rAux["de_empr_distdora_nombre"] != System.DBNull.Value)
                                            Distribuidora = rAux["de_empr_distdora_nombre"].ToString();
                                    }
                                    dbAux.CloseConnection();
                                }
                                //// Fin mirar en el historico

                                //////if (r["tx_apell_cli"] != System.DBNull.Value)
                                //////    workSheet.Cells[f, c].Value = r["tx_apell_cli"].ToString();
                                //////c++;
                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = apellido;
                                }
                                else
                                {
                                    if (r["tx_apell_cli"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["tx_apell_cli"].ToString();
                                }
                                c++;


                                //DIAS_ESTADO_SAP
                                workSheet.Cells[f, c].Value = GetDiasEstado(r["cups20"].ToString());
                                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                c++;

                                //DIAS_ESTADO_KEE
                                workSheet.Cells[f, c].Value = "";c++;

                                //DIAS_ESTADO_GLOBAL
                                workSheet.Cells[f, c].Value = ""; c++;

                                //Recuperar Incidencia Facturacion
                                listaIncidenciasSAP = RelacionIncidenciasSAP.GetCups(r["cups20"].ToString() + aniomes);

                                if (listaIncidenciasSAP.Count == 0)
                                {
                                    IncidenciaFacturacion = "";
                                    Estado_FAC_SE = "";
                                    Titulo_FAC = "";
                                }
                                else
                                {
                                    foreach (var mm in listaIncidenciasSAP)
                                    {
                                        string[] cadena = mm.Split(';');
                                        IncidenciaFacturacion = cadena[0];
                                        Estado_FAC_SE = cadena[1];
                                        Titulo_FAC = cadena[2];
                                        break;
                                    }
                                }
                                //Fin recuperar incidencia facturacion

                                // Comprobamos con la tabla Relacion_INC_CUPS
                                ////// 1.Se relaciona el detalle del informe obtenido(campos cups20 y mes) con la tabla "Relacion_INC_CUPS"(campos cups y mes).Si para un cups se encuentra en esa relación una incidencia que afeca al mismo mes que está pendiente se deberá desechar el estado que tuviera el informe y actualizarlo por uno nuevo.
                                ////// 2.El estado a actualizar dependerá del campo "area" de la tabla "Relacion_INC_CUPS", pudiendo existir 3 opciones:
                                ////// --01.B09 Incidencia_Contratacion: se informará con este estado cuando encuentre para ese mes una incidencia en el area de contratacion, independientemente de que también haya otras para ese mismo mes en otras areas.
                                //////-- 01.B10 Incidencia_Medida: no estando en el caso anterior, se informará con este estado cuando encuentre para ese mes una incidencia en el area de medida, independientemente de que también haya otras para ese mismo mes en el area de facturacion.
                                ////// --01.B11 Incidencia_Facturacion: no estando en ninguno de los dos casos anteriores, se informará con este estado cuando encuentre para ese mes una incidencia en el area de Facturacion.
                                strSql = " select area, incidencia, estado_incidencia, fecha_apertura, prioridad_negocio, titulo, e_s_estado, Reincidente from Relacion_INC_CUPS"
                                         + " where cups='" + r["cups20"].ToString() + "' and Mes_pendiente='" + aniomes + "'"
                                         + " order by Fecha_apertura asc";

                                dbIncidencia = new MySQLDB(MySQLDB.Esquemas.MED);
                                command = new MySqlCommand(strSql, dbIncidencia.con);
                                inci = command.ExecuteReader();

                                areaincidencia = "";
                                incidencia = "";
                                estado_incidencia = "";
                                fecha_apertura = "";
                                prioridad_negocio = "";
                                titulo = "";
                                e_s_estado = "";
                                subestado_incidencia = "";
                                reincidente = "";
                                ExisteIncidencia = false;

                                while (inci.Read())
                                {
                                    ExisteIncidencia = true;
                                    if (areaincidencia == "")
                                    {
                                        areaincidencia = inci["area"].ToString();

                                        if (areaincidencia == "Contratacion")
                                            subestado_incidencia = "01.B09 Incidencia_Contratacion";
                                        if (areaincidencia == "Medida")
                                            subestado_incidencia = "01.B10 Incidencia_Medida";
                                        if (areaincidencia == "Facturacion")
                                            subestado_incidencia = "01.B11 Incidencia_Facturacion";

                                        incidencia = inci["incidencia"].ToString();
                                        estado_incidencia = inci["estado_incidencia"].ToString();
                                        fecha_apertura = inci["fecha_apertura"].ToString();
                                        prioridad_negocio = inci["prioridad_negocio"].ToString();
                                        titulo = inci["titulo"].ToString();
                                        e_s_estado = inci["e_s_estado"].ToString();
                                        reincidente = inci["Reincidente"].ToString();
                                    }
                                    else
                                    {
                                        if (areaincidencia != inci["area"].ToString())
                                        {
                                            areaincidencia = inci["area"].ToString();

                                            if (areaincidencia == "Contratacion")
                                                subestado_incidencia = subestado_incidencia + " - 01.B09 Incidencia_Contratacion";
                                            if (areaincidencia == "Medida")
                                                subestado_incidencia = subestado_incidencia + " - 01.B10 Incidencia_Medida";
                                            if (areaincidencia == "Facturacion")
                                                subestado_incidencia = subestado_incidencia + " - 01.B11 Incidencia_Facturacion";

                                            incidencia = incidencia + " - " + inci["incidencia"].ToString();
                                            estado_incidencia = estado_incidencia + " - " + inci["estado_incidencia"].ToString();
                                            fecha_apertura = fecha_apertura + " - " + inci["fecha_apertura"].ToString();
                                            prioridad_negocio = prioridad_negocio + " - " + inci["prioridad_negocio"].ToString();
                                            titulo = titulo + " - " + inci["titulo"].ToString();
                                            e_s_estado = e_s_estado + " - " + inci["e_s_estado"].ToString();
                                        }
                                    }
                                } // Fin  while (inci.Read())

                                dbIncidencia.CloseConnection();
                                if (ExisteIncidencia == true)
                                {
                                    c--;
                                    workSheet.Cells[f, c].Value = subestado_incidencia; c++;
                                    workSheet.Cells[f, c].Value = IncidenciaFacturacion; c++;
                                    //////workSheet.Cells[f, c].Value = subestado_incidencia; c++;

                                    workSheet.Cells[f, c].Value = Estado_FAC_SE; c++;
                                    workSheet.Cells[f, c].Value = Titulo_FAC; c++;
                                    workSheet.Cells[f, c].Value = incidencia; c++;
                                    workSheet.Cells[f, c].Value = "";c++;
                                    workSheet.Cells[f, c].Value = estado_incidencia; c++;
                                    workSheet.Cells[f, c].Value = fecha_apertura; c++;
                                    workSheet.Cells[f, c].Value = prioridad_negocio; c++;
                                    workSheet.Cells[f, c].Value = titulo; c++;
                                    workSheet.Cells[f, c].Value = e_s_estado; c++;
                                }
                                else
                                {
                                    workSheet.Cells[f, c].Value = IncidenciaFacturacion; c++;
                                    workSheet.Cells[f, c].Value = Estado_FAC_SE; c++;
                                    workSheet.Cells[f, c].Value = Titulo_FAC; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                         
                                }
                                ////////////////////////////////////////////////////////////////////////////////////////////////

                                ///FH_ALTA_SALESFORCE
                                //////workSheet.Cells[f, c].Value = ""; c++;

                                if (r["fh_alta_crto"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_alta_crto"]).Date;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                else
                                {

                                }
                                c++;
                                // FIN FH_ALTA_SALESFORCE

                                //FECHA_ALTA_KEE*****************************************************
                                listaAltaKEE = pendienteWeb_B2B.GetCupsFechaAltaKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fec_act"]));

                                if (listaAltaKEE.Count == 0)
                                {
                                    workSheet.Cells[f, c].Value = ""; c++;
                                }
                                else
                                {
                                    if (listaAltaKEE[0] == "01/01/0001 0:00:00")
                                    {
                                        workSheet.Cells[f, c].Value = ""; c++;
                                    }
                                    else
                                    {

                                        foreach (var mm in listaAltaKEE)
                                        {
                                            workSheet.Cells[f, c].Value = mm; c++;
                                            break;
                                        }
                                        //////workSheet.Cells[f, c].Value = lista; c++;
                                    }
                                }

                                //FIN FH_ALTA_KEE*****************************************************

                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToDateTime(FechaAlta).Date;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                else
                                {
                                    if (r["fh_alta_crto"] != System.DBNull.Value)
                                    {
                                        workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_alta_crto"]).Date;
                                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    }
                                }
                                c++;



                                if (r["fh_baja_crto"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_baja_crto"]).Date;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                                }
                                else
                                {
                                    ////////if (r["fh_prev_fin_crto"] != System.DBNull.Value)
                                    ////////{
                                    ////////    workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_prev_fin_crto"]).Date;
                                    ////////    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                                    ////////}
                                }
                                c++;

                                //Paco, buscamos la fecha de Baja de SAP ******************************************
                                CadenaAuxiliar = r["id_crto_ext"].ToString();
                                CadenaAuxiliar = CadenaAuxiliar.PadLeft(12, '0'); //Relleno con ceros a la izquierda hasta 12


                                fhBajaSap = FechaBajaSap.GetCupsFechaBajaSAP(CadenaAuxiliar);

                                if (fhBajaSap.Count == 0)
                                {
                                    workSheet.Cells[f, c].Value = ""; c++;
                                }
                                else
                                {
                                    foreach (var mm in fhBajaSap)
                                    {
                                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                        workSheet.Cells[f, c].Value = mm; c++; 
                                        break;
                                    }
                                }

                                //////strSql = "SELECT  max(cd_sec_crto), fh_baja from ed_owner.t_ed_h_sap_crto_front "
                                //////+ " WHERE id_crto_ext = '" + CadenaAuxiliar + "'"
                                //////+ " group by  fh_baja"
                                //////+ " order by  max(cd_sec_crto) desc";

                                //////dbRS = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                                //////commandRS = new OdbcCommand(strSql, dbRS.con);
                                //////rRS = commandRS.ExecuteReader();
                                //////while (rRS.Read())
                                //////{
                                //////    if (rRS["fh_baja"] != System.DBNull.Value )
                                //////    {
                                //////        if (rRS["fh_baja"].ToString() != "01/01/1400 0:00:00")
                                //////        {
                                //////            workSheet.Cells[f, c].Value = Convert.ToDateTime(rRS["fh_baja"]).Date;
                                //////            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                //////        }
                                //////    }
                                //////    break;
                                //////}
                                //////dbRS.CloseConnection();
                                ///
                                //////c++;

                                //Fin - Fh_baja de SAP ***************************************************

                                //////if (r["cd_empr"] != System.DBNull.Value)
                                //////    workSheet.Cells[f, c].Value = r["cd_empr"].ToString();
                                //////c++;

                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = empresa;
                                }
                                else
                                {
                                    if (r["cd_empr"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["cd_empr"].ToString();
                                }
                                c++;

                                if (r["cd_segmento_ptg_Con_compor"] != System.DBNull.Value) //SEGMENTO
                                    workSheet.Cells[f, c].Value = r["cd_segmento_ptg_Con_compor"].ToString();
                                c++;

                                //////if (r["de_tp_cli"] != System.DBNull.Value)
                                //////    workSheet.Cells[f, c].Value = r["de_tp_cli"].ToString();
                                //////c++;

                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = cliente;
                                }
                                else
                                {
                                    if (r["de_tp_cli"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["de_tp_cli"].ToString();
                                }
                                c++;

                                //////if (r["cd_nif_cif_cli"] != System.DBNull.Value)
                                //////    workSheet.Cells[f, c].Value = r["cd_nif_cif_cli"].ToString();
                                //////c++;

                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = nif;
                                }
                                else
                                {
                                    if (r["cd_nif_cif_cli"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["cd_nif_cif_cli"].ToString();
                                }
                                c++;

                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToDateTime(FechaInicio).Date;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                else
                                {
                                    if (r["fh_inicio_vers_crto"] != System.DBNull.Value)
                                    {
                                        workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_inicio_vers_crto"]).Date;
                                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    }
                                }
                                c++;

                                if (r["id_instalacion"] != System.DBNull.Value)
                                    workSheet.Cells[f, c].Value = r["id_instalacion"].ToString();
                                c++;
                                //////if (r["cd_tarifa_c"] != System.DBNull.Value)
                                //////    workSheet.Cells[f, c].Value = r["cd_tarifa_c"].ToString();
                                //////c++;

                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = Tarifa;
                                }
                                else
                                {
                                    if (r["cd_tarifa_c"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["cd_tarifa_c"].ToString();
                                }
                                c++;

                                //////if (r["cd_crto_comercial"] != System.DBNull.Value)
                                //////    workSheet.Cells[f, c].Value = r["cd_crto_comercial"].ToString();
                                //////c++;

                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = contrato;
                                }
                                else
                                {
                                    if (r["cd_crto_comercial"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["cd_crto_comercial"].ToString();
                                }
                                c++;

                                //////if (r["de_empr_distdora_nombre"] != System.DBNull.Value)
                                //////    workSheet.Cells[f, c].Value = r["de_empr_distdora_nombre"].ToString();
                                //////c++;

                                if (Existetedhps == false)
                                {
                                    workSheet.Cells[f, c].Value = Distribuidora;
                                }
                                else
                                {
                                    if (r["de_empr_distdora_nombre"] != System.DBNull.Value)
                                        workSheet.Cells[f, c].Value = r["de_empr_distdora_nombre"].ToString();
                                }
                                c++;

                                if (r["de_estado"] != System.DBNull.Value)
                                    workSheet.Cells[f, c].Value = r["de_estado"].ToString();
                                c++;
                                if (r["de_subestado"] != System.DBNull.Value)
                                    workSheet.Cells[f, c].Value = r["de_subestado"].ToString();
                                c++;

                                //if (r["lg_multimedida"] != System.DBNull.Value)
                                //{
                                //    workSheet.Cells[f, c].Value = r["lg_multimedida"].ToString();

                                //}
                                //else
                                //{
                                //    workSheet.Cells[f, c].Value = "N";
                                //}

                                //////workSheet.Cells[f, c].Value = GetDiasEstado(r["cups20"].ToString());
                                //////workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //////c++;

                                if (r["TAM"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                }
                                c++;

                                // Paco - añadimos la ultima fecha desde y hasta facturada t_ed_h_sap_facts
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                                List<string> listaInicio = new List<string>();
                                listaInicio = pendienteSAPKEE.GetCupsInicioUltimaFacturada(r["cups20"].ToString());
                                if (listaInicio != null)
                                {
                                    if (listaInicio.Count == 0)
                                    {
                                        workSheet.Cells[f, c].Value = ""; c++;
                                    }
                                    else
                                    {
                                        if (listaInicio[0] == "01/01/0001 0:00:00")
                                        {
                                            workSheet.Cells[f, c].Value = ""; c++;
                                        }
                                        else
                                        {

                                            foreach (var mm in listaInicio)
                                            {
                                                workSheet.Cells[f, c].Value = mm; c++;
                                                break;
                                            }
                                            ////workSheet.Cells[f, c].Value = lista; c++;
                                        }
                                    }
                                }
                                else
                                {
                                    workSheet.Cells[f, c].Value = ""; c++;
                                }

                                // Paco - añadimos la ultima fecha desde y hasta facturada t_ed_h_sap_facts
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                                List<string> listaFin = new List<string>();
                                listaFin = pendienteSAPKEE.GetCupsFinalUltimaFacturada(r["cups20"].ToString());
                                if (listaFin != null)
                                {
                                    if (listaFin.Count == 0)
                                    {
                                        workSheet.Cells[f, c].Value = ""; c++;
                                    }
                                    else
                                    {
                                        if (listaFin[0] == "01/01/0001 0:00:00")
                                        {
                                            workSheet.Cells[f, c].Value = ""; c++;
                                        }
                                        else
                                        {

                                            foreach (var mm in listaFin)
                                            {
                                                workSheet.Cells[f, c].Value = mm; c++;
                                                break;
                                            }
                                            ////workSheet.Cells[f, c].Value = lista; c++;
                                        }
                                    }
                                }
                                else
                                {
                                    workSheet.Cells[f, c].Value = ""; c++;
                                }


                                subestadosKronosBTN.existe = false;

                                if (subestados_sap.AreaResponsableMedida(r["cd_subestado"].ToString()))
                                {
                                    subestadosKronosBTN.existe = false;


                                    //IMPORTANTE!!!! Limpiamos el area para que no queda una antigua en la revisión de incidencias
                                    subestadosKronosBTN.area_responsable = null;

                                    //Paco -- He duplicado funciones para marcar los casos que no hay datos en kronos para la fecha GetEstadoKEEDetalle y GetCupsDetalle
                                    //Faltaria controlar los casos en los que vienen segmentos, es decir, para el mes que se mira de Sap vienen dos segmentos
                                    //en Kronos, habria que quedarse con el de fecha más antigua, he puesto codigo sin probar para este caso en GetCupsDetalle
                                    //Lo que hago es crear a pelo un estado "Discrepancia: ..... " que es la etiqueta que pongo y que no está en la tabla parametrica estados_kee_param 
                                    subestadosKronosBTN.GetEstadoKEEDetalleBTN(pendienteWeb_B2BBTN.GetCupsDetalle_BTN(r["cups20"].ToString(),
                                        Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]), Convert.ToDateTime(r["fec_act"])));

                                    //////if (subestadosKronos.multimedida == true)
                                    //////{
                                    //////    Debug.Print(r["cups20"].ToString());
                                    //////}
                            
                                    if (subestadosKronosBTN.existe)
                                    {
                                        BuscaCadena = subestadosKronosBTN.descripcion_estado;
                                        List<string> lista = new List<string>();

                                        if (BuscaCadena.IndexOf("Discrepancia") >= 0)
                                        {
                                            if (BuscaCadena.IndexOf("No existe el cups en el informe del pendiente de KEE") >= 0)
                                            {
                                                //Paco 30/05/2024 Se trata de manera diferente a ES-MTBTE, son puntos migrados el pasado fin de semana que no tienen todavía medidas gestionadas 
                                                //directamente en KEE. Como las medidas migradas no se cargan en BI, no localizas lecturas. 
                                                NumeroDias = 0;
                                                NumeroDias = (DateTime.Now - Convert.ToDateTime(r["fh_desde"])).Days;

                                                //Controlo si la fecha desde -1 hasta el día de hoy supera los 90 días
                                                if (NumeroDias > 45)
                                                {
                                                    workSheet.Cells[f, c].Value = "Pendiente KEE"; c++;
                                                    workSheet.Cells[f, c].Value = "MEDIDA"; c++;
                                                    workSheet.Cells[f, c].Value = "01.C25 Pendiente Sistemas KEE - Pte ejecución algoritmo"; c++;
                                                    workSheet.Cells[f, c].Value = "Pendiente Sistemas KEE - Pte ejecución algoritmo"; c++;
                                                }
                                                else
                                                {
                                                    workSheet.Cells[f, c].Value = "Pendiente KEE"; c++;
                                                    workSheet.Cells[f, c].Value = "DISTRIBUIDORA"; c++;
                                                    workSheet.Cells[f, c].Value = "01.C25 Pendiente Medida KEE - En ciclo de Medida"; c++;
                                                    workSheet.Cells[f, c].Value = "Pendiente Medida KEE - En ciclo de Medida"; c++;
                                                }

                                                //Fin /Paco 30/05/2024

                                                //OLD
                                                //////workSheet.Cells[f, c].Value = ""; c++;
                                                //////workSheet.Cells[f, c].Value = ""; c++;
                                                //////workSheet.Cells[f, c].Value = "01.B07 Error Sistemas KEE-SAP - Pdte SAP-No existe en KEE"; c++;
                                                //////workSheet.Cells[f, c].Value = ""; c++;
                                                //FIN OLD

                                                workSheet.Cells[f, c].Value = ""; c++; //FH_DESDE_KEE
                                                workSheet.Cells[f, c].Value = ""; c++; //FH_HASTA_KEE
                                                workSheet.Cells[f, c].Value = ""; c++;
                                                workSheet.Cells[f, c].Value = subestadosKronosBTN.descripcion_estado; ; c++;

                                                //Subestado Global
                                                if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                                {
                                                    workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoKEE].Value; c++;
                                                }
                                                else
                                                {
                                                    workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoSAP].Value; c++;
                                                }
                                                //Fin Subestado Global


                                                //Paco 30/05/2024 
                                                if (NumeroDias > 45)
                                                {
                                                    workSheet.Cells[f, c].Value = "02_INCIDENCIA_MEDIDA"; c++;
                                                    workSheet.Cells[f, c].Value = "02_INCIDENCIA"; c++;
                                                }
                                                else
                                                {
                                                    workSheet.Cells[f, c].Value = "01_DISTRIBUIDORA"; c++;
                                                    workSheet.Cells[f, c].Value = "01_DISTRIBUIDORA"; c++;
                                                }
                                                //Fin Paco 30/05/2024 

                                                //////workSheet.Cells[f, c].Value = ""; c++;
                                                //////workSheet.Cells[f, c].Value = ""; c++;
                                                ///

                                                //AREA
                                                workSheet.Cells[f, c].Value = ""; c++;
                                                //•	Nº DÍAS ESTADO KEE y •	Nº DÍAS PTE KEE
                                                workSheet.Cells[f, PosicionDiasEstadoKEE].Value = ""; 
                                                workSheet.Cells[f, c].Value = ""; c++;
                                            }
                                            else
                                            {
                                                ///Discrepancia: No existen fechas en el informe pendiente de KEE para el periodo
                                                if (BuscaCadena.IndexOf("No existen fechas en el informe pendiente de KEE para el periodo") >= 0)
                                                {
                                                    NumeroDias = 0;

                                                    string[] cadena = subestadosKronosBTN.descripcion_estado.Split(';');
                                                    string[] cadenaAux = cadena[1].Split('/');

                                                    //////802 Facturada DISTRIBUIDORA       Pendiente KEE   01.C25  Pendiente Medida KEE - En ciclo de Medida           01_DISTRIBUIDORA      01_DISTRIBUIDORA
                                                    //////802 Facturada MEDIDA              Pendiente KEE   01.C25  Pendiente Sistemas KEE - Pte ejecución algoritmo    02_INCIDENCIA_MEDIDA  02_INCIDENCIA

                                                    fecha_calculada = Convert.ToDateTime( cadenaAux[0].Substring(1, 8).Substring(6, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(4, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(0, 4));
                                                    fecha_calculada = fecha_calculada.AddDays(-1);
                                                    NumeroDias =(DateTime.Now - fecha_calculada).Days;

                                                    //Controlo si la fecha desde -1 hasta el día de hoy supera los 90 días
                                                    if (NumeroDias > 45)
                                                    {
                                                        workSheet.Cells[f, c].Value = "Pendiente KEE"; c++;
                                                        workSheet.Cells[f, c].Value = "MEDIDA"; c++;
                                                        workSheet.Cells[f, c].Value = "01.C25 Pendiente Sistemas KEE - Pte ejecución algoritmo"; c++;
                                                        workSheet.Cells[f, c].Value = "Pendiente Sistemas KEE - Pte ejecución algoritmo"; c++;
                                                    }
                                                    else
                                                    {
                                                        workSheet.Cells[f, c].Value = "Pendiente KEE"; c++;
                                                        workSheet.Cells[f, c].Value = "DISTRIBUIDORA"; c++;
                                                        workSheet.Cells[f, c].Value = "01.C25 Pendiente Medida KEE - En ciclo de Medida"; c++;
                                                        workSheet.Cells[f, c].Value = "Pendiente Medida KEE - En ciclo de Medida"; c++;
                                                    }

                                                    workSheet.Cells[f, c].Value = ""; c++; //FH_DESDE_KEE
                                                    workSheet.Cells[f, c].Value = ""; c++; //FH_HASTA_KEE
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = "Nº días en ciclo de medida: " + NumeroDias.ToString (); c++;

                                                    //Subestado Global
                                                    if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                                    {
                                                        workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoKEE].Value; c++;
                                                    }
                                                    else
                                                    {
                                                        workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoSAP].Value; c++;
                                                    }
                                                    //Fin Subestado Global

                                                    if (NumeroDias > 45)
                                                    {
                                                        workSheet.Cells[f, c].Value = "02_INCIDENCIA_MEDIDA"; c++;
                                                        workSheet.Cells[f, c].Value = "02_INCIDENCIA"; c++;
                                                    }
                                                    else
                                                    {
                                                        workSheet.Cells[f, c].Value = "01_DISTRIBUIDORA"; c++;
                                                        workSheet.Cells[f, c].Value = "01_DISTRIBUIDORA"; c++;
                                                    }

                                                    //AREA
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    //•	Nº DÍAS ESTADO KEE y •	Nº DÍAS PTE KEE
                                                    workSheet.Cells[f, PosicionDiasEstadoKEE].Value = ""; 
                                                    workSheet.Cells[f, c].Value = ""; c++;

                                                }
                                                else
                                                {
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;

                                                    //////string[] cadena = subestadosKronosBTN.descripcion_estado.Split(';');
                                                    //////string[] cadenaAux = cadena[1].Split('/');
                                                    //////workSheet.Cells[f, c].Value = cadenaAux[0].Substring(1, 8).Substring(6, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(4, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(0, 4); c++;
                                                    //////workSheet.Cells[f, c].Value = cadenaAux[1].Substring(0, 8).Substring(6, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(4, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(0, 4); c++;

                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;

                                                    ///lista = pendienteWeb_B2B.GetCupsFinKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]));
                                                    lista = pendienteWeb_B2B.GetCupsFinKEENew(r["cups20"].ToString(), Convert.ToDateTime(r["fec_act"]));

                                                    if (lista.Count == 0)
                                                    {
                                                        workSheet.Cells[f, c].Value = ""; c++;
                                                    }
                                                    else
                                                    {
                                                        if (lista[0] == "01/01/0001 0:00:00")
                                                        {
                                                            workSheet.Cells[f, c].Value = ""; c++;
                                                        }
                                                        else
                                                        {

                                                            foreach (var mm in lista)
                                                            {
                                                                //Lo quito hasta que diga Adriana
                                                                workSheet.Cells[f, c].Value = mm;
                                                                c++;
                                                                break;
                                                            }
                                                            //////workSheet.Cells[f, c].Value = lista; c++;
                                                        }
                                                    }
                                                
                                                    workSheet.Cells[f, c].Value = subestadosKronosBTN.descripcion_estado;
                                                    c++;

                                                    //Subestado Global
                                                    if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                                    {
                                                        workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoKEE].Value; c++;
                                                    }
                                                    else
                                                    {
                                                        workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoSAP].Value; c++;
                                                    }

                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                    //•	Nº DÍAS ESTADO KEE y •	Nº DÍAS PTE KEE
                                                    workSheet.Cells[f, PosicionDiasEstadoKEE].Value = ""; 
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            //Lo quito hasta que diga Adriana
                                            ////////workSheet.Cells[f, c].Value = subestadosKronosBTN.estado_periodo; c++;

                                            ////////////////Control del algoritmo para el estado 802
                                            //////////////if (subestadosKronosBTN.cod_estado == "802")
                                            //////////////{
                                            //////////////    if ((DateTime.Now - subestadosKronosBTN.fh_hasta).Days > 90)
                                            //////////////    {
                                            //////////////        workSheet.Cells[f, c].Value = "MEDIDA"; c++;
                                            //////////////        workSheet.Cells[f, c].Value = "01.C25 Pendiente Sistemas KEE - Pte ejecución algoritmo"; c++;
                                            //////////////    }
                                            //////////////    else
                                            //////////////    {
                                            //////////////        workSheet.Cells[f, c].Value = subestadosKronosBTN.area_responsable; c++;
                                            //////////////        workSheet.Cells[f, c].Value = subestadosKronosBTN.descripcion_subestado; c++;
                                            //////////////    }
                                            //////////////}
                                            //////////////else
                                            //////////////{
                                            ////////    workSheet.Cells[f, c].Value = subestadosKronosBTN.area_responsable; c++;
                                            ////////    workSheet.Cells[f, c].Value = subestadosKronosBTN.descripcion_subestado; c++;
                                            //////////////}

                                            ////////workSheet.Cells[f, c].Value = subestadosKronosBTN.descripcion_estado; c++;


                                            //04-04-2025
                                            if (subestadosKronosBTN.cod_estado == "802")
                                            {
                                                NumeroDias = (DateTime.Now - subestadosKronosBTN.fh_hasta).Days;

                                               //////Controlo si la fecha desde - 1 hasta el día de hoy supera los 45 días
                                                if (NumeroDias > 45)
                                                {
                                                    workSheet.Cells[f, c].Value = "Pendiente KEE"; c++;
                                                    workSheet.Cells[f, c].Value = "MEDIDA"; c++;
                                                    workSheet.Cells[f, c].Value = "01.C25 Pendiente Sistemas KEE - Pte ejecución algoritmo"; c++;
                                                    workSheet.Cells[f, c].Value = "Pendiente Sistemas KEE - Pte ejecución algoritmo"; c++;
                                                }
                                                else
                                                {
                                                    workSheet.Cells[f, c].Value = "Pendiente KEE"; c++;
                                                    workSheet.Cells[f, c].Value = "DISTRIBUIDORA"; c++;
                                                    workSheet.Cells[f, c].Value = "01.C25 Pendiente Medida KEE - En ciclo de Medida"; c++;
                                                    workSheet.Cells[f, c].Value = "Pendiente Medida KEE - En ciclo de Medida"; c++;
                                                }
                                            }
                                            else
                                            {

                                                workSheet.Cells[f, c].Value = subestadosKronosBTN.estado_periodo; c++;
                                                workSheet.Cells[f, c].Value = subestadosKronosBTN.area_responsable; c++;
                                                workSheet.Cells[f, c].Value = subestadosKronosBTN.descripcion_subestado; c++;
                                                workSheet.Cells[f, c].Value = subestadosKronosBTN.descripcion_estado; c++;
                                            }

                                            //Fin 04-04-2025

                                            // fecha_desde_kee y fecha_hasta_kee
                                            workSheet.Cells[f, c].Value = subestadosKronosBTN.fh_desde;
                                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                            c++;
                                            workSheet.Cells[f, c].Value = subestadosKronosBTN.fh_hasta;
                                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                            c++;
                                            //////lista = pendienteWeb_B2B.GetCupsFinKEE(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]), Convert.ToDateTime(r["fh_hasta"]));
                                            lista = pendienteWeb_B2B.GetCupsFinKEENew(r["cups20"].ToString(), Convert.ToDateTime(r["fec_act"]));

                                            if (lista.Count == 0)
                                            {
                                                workSheet.Cells[f, c].Value = ""; c++;
                                            }
                                            else
                                            {
                                                if (lista[0] == "01/01/0001 0:00:00")
                                                {
                                                    workSheet.Cells[f, c].Value = ""; c++;
                                                }
                                                else
                                                {

                                                    foreach (var mm in lista)
                                                    {
                                                        workSheet.Cells[f, c].Value = mm; c++;
                                                        break;
                                                    }
                                                    //////workSheet.Cells[f, c].Value = lista; c++;
                                                }
                                            }

                                            workSheet.Cells[f, c].Value = ""; c++;

                                            //Subestado Global
                                            if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                            {
                                                workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoKEE].Value; c++;
                                            }
                                            else
                                            {
                                                workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoSAP].Value; c++;
                                            }

                                            ////////Control del algoritmo para el estado 802
                                            //////if (subestadosKronosBTN.cod_estado == "802")
                                            //////{
                                            //////    if ((DateTime.Now - subestadosKronosBTN.fh_hasta).Days > 90) {
                                            //////        workSheet.Cells[f, c].Value = "02_INCIDENCIA_MEDIDA"; c++;
                                            //////        workSheet.Cells[f, c].Value = "02_INCIDENCIA"; c++;
                                            //////    }
                                            //////    else {
                                            //////        workSheet.Cells[f, c].Value = subestadosKronosBTN.ESTADO_GLOBAL; c++;
                                            //////        workSheet.Cells[f, c].Value = subestadosKronosBTN.ESTADO_GLOBAL_A_REPORTAR; c++;
                                            //////    }
                                            //////}
                                            //////else
                                            //////{
                                            ///

                                            //Paco 18/12/2024 OLD
                                            workSheet.Cells[f, c].Value = subestadosKronosBTN.ESTADO_GLOBAL; c++;
                                            workSheet.Cells[f, c].Value = subestadosKronosBTN.ESTADO_GLOBAL_A_REPORTAR; c++;
                                            //Fin Paco 18/12/2024 OLD


                                            //Paco 18/12/2024
                                            //////string[] cadena = subestadosKronosBTN.descripcion_estado.Split(';');
                                            //////string[] cadenaAux = cadena[1].Split('/');

                                            //////fecha_calculada = Convert.ToDateTime(cadenaAux[0].Substring(1, 8).Substring(6, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(4, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(0, 4));
                                            //////fecha_calculada = fecha_calculada.AddDays(-1);
                                            //////NumeroDias = (DateTime.Now - fecha_calculada).Days;

                                            //////if (NumeroDias > 45)
                                            //////{
                                            //////    workSheet.Cells[f, c].Value = "02_INCIDENCIA_MEDIDA"; c++;
                                            //////    workSheet.Cells[f, c].Value = "02_INCIDENCIA"; c++;
                                            //////}
                                            //////else
                                            //////{
                                            //////    workSheet.Cells[f, c].Value = "01_DISTRIBUIDORA"; c++;
                                            //////    workSheet.Cells[f, c].Value = "01_DISTRIBUIDORA"; c++;
                                            //////}
                                            //Paco 18/12/2024


                                            //'AREA'
                                            workSheet.Cells[f, c].Value = ""; c++;
                                            //////}

                                            //////workSheet.Cells[f, c].Value= subestadosKronosBTN.ESTADO_GLOBAL; c++;
                                            //////workSheet.Cells[f, c].Value = subestadosKronosBTN.ESTADO_GLOBAL_A_REPORTAR; c++;


                                            //•	Nº DÍAS ESTADO KEE y •	Nº DÍAS PTE KEE
                                            if (subestadosKronosBTN.fecha_modificacion.ToString() == "01/01/0001 0:00:00") {
                                                workSheet.Cells[f, PosicionDiasEstadoKEE].Value = 1; 
                                                NumeroDiasKEE = 1;
                                            }
                                            else {
                                                workSheet.Cells[f, PosicionDiasEstadoKEE].Value = (DateTime.Now - subestadosKronosBTN.fecha_modificacion).Days; 
                                                NumeroDiasKEE = (DateTime.Now - subestadosKronosBTN.fecha_modificacion).Days;
                                            }
                                            workSheet.Cells[f, c].Value = (DateTime.Now - subestadosKronosBTN.fh_hasta).Days; c++;
                                        }
                                    }   // Fin If existe
                                }

                                else
                                {
                                    //Para que llegue hasta la última columna aunque no exista en Kronos
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    //////workSheet.Cells[f, c].Value = ""; c++;
                                    ///
                                    List<string> listaAux = new List<string>();
                                    listaAux = pendienteWeb_B2B.GetCupsFinKEENew(r["cups20"].ToString(), Convert.ToDateTime(r["fec_act"]));

                                    if (listaAux.Count == 0)
                                    {
                                        workSheet.Cells[f, c].Value = ""; c++;
                                    }
                                    else
                                    {
                                        if (listaAux[0] == "01/01/0001 0:00:00")
                                        {
                                            workSheet.Cells[f, c].Value = ""; c++;
                                        }
                                        else
                                        {

                                            foreach (var mm in listaAux)
                                            {
                                                //Lo quito hasta que diga Adriana
                                                workSheet.Cells[f, c].Value = mm;
                                                c++;
                                                break;
                                            }
                                            //////workSheet.Cells[f, c].Value = lista; c++;
                                        }
                                    }

                                    //Este es la discrepancia
                                    workSheet.Cells[f, c].Value = ""; c++;

                                    //Subestado Global
                                    if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                    {
                                        workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoKEE].Value; c++;
                                    }
                                    else
                                    {
                                        workSheet.Cells[f, c].Value = workSheet.Cells[f, PosicionSubEstadoSAP].Value; c++;
                                    }
                                    //ESTADO GLOBAL Y ESTADO GLOBAL A REPORTAR
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    workSheet.Cells[f, c].Value = ""; c++;
                                    //•	Nº DÍAS ESTADO KEE y •	Nº DÍAS PTE KEE
                                    workSheet.Cells[f, PosicionDiasEstadoKEE].Value = ""; 
                                    workSheet.Cells[f, c].Value = ""; c++;
                                }
                                System.Diagnostics.Debug.WriteLine(r["cups20"].ToString() + "-" + r["de_subestado"].ToString() + "-" + subestadosKronos.descripcion_subestado + "-" + subestadosKronos.descripcion_estado);

                                //Paco 09/05/2024 **************************************************************************************
                                //Si se ha encontrado una  incidencia en Relacion_INC_CUPS (Tiene un AREA MARCADA) --> Si la incidencia tiene la misma AREA que la que aparece en la tabla ESTADOS_KEE_PARAM (subestadosKronos.area_responsable)
                                // areaincidencia=  subestadosKronos.area_responsable
                                // EN BTN SITENGO INCIDENCIA actualizo EL ESTADO GLOBAL Y EL ESTADO GLOBAL A REPORTAR , PERO NO CAMBIO SUBESTADOGLOBal

                                ExisteIncidenciaBTN = false;
                                if (areaincidencia != "" & subestadosKronosBTN.existe==true)
                                {
                                    Auxiliar = Reincidente(areaincidencia, subestadosKronosBTN.area_responsable);
                                    if (Auxiliar != "")
                                    {
                                        if (Auxiliar == "01.B09 Incidencia_Contratacion")
                                        {
                                            workSheet.Cells[f, PosicionSubEstadoGlobal].Value = "01.B09 Incidencia_Contratacion";
                                            workSheet.Cells[f, PosicionSubEstadoGlobal + 1].Value = "02_INCIDENCIA_CONTRATACION";
                                            workSheet.Cells[f, PosicionSubEstadoGlobal + 2].Value = "02_INCIDENCIA";
                                            workSheet.Cells[f, PosicionReincidente].Value = reincidente;
                                            ExisteIncidenciaBTN = true;
                                        }
                                        if (Auxiliar == "01.B10 Incidencia_Medida")
                                        {
                                            workSheet.Cells[f, PosicionSubEstadoGlobal].Value = "01.B10 Incidencia_Medida";
                                            workSheet.Cells[f, PosicionSubEstadoGlobal + 1].Value = "02_INCIDENCIA_MEDIDA";
                                            workSheet.Cells[f, PosicionSubEstadoGlobal + 2].Value = "02_INCIDENCIA";
                                            workSheet.Cells[f, PosicionReincidente].Value = reincidente;
                                            ExisteIncidenciaBTN = true;
                                        }
                                        if (Auxiliar == "01.B11 Incidencia_Facturacion")
                                        {
                                            workSheet.Cells[f, PosicionSubEstadoGlobal].Value = "01.B11 Incidencia_Facturacion";
                                            workSheet.Cells[f, PosicionSubEstadoGlobal + 1].Value = "02_INCIDENCIA_FACTURACION";
                                            workSheet.Cells[f, PosicionSubEstadoGlobal + 2].Value = "02_INCIDENCIA";
                                            workSheet.Cells[f, PosicionReincidente].Value = reincidente;
                                            ExisteIncidenciaBTN = true;
                                        }
                                    }

                                    //////else
                                    //////{ //No coinciden las areas de incidencia y area responsable o el arearesponsable no está rellena
                                    //////    if (ExisteIncidencia == true)
                                    //////    {
                                    //////        if (areaincidencia == "Medida")
                                    //////        {
                                    //////            workSheet.Cells[f, PosicionSubEstadoGlobal].Value = "01.B10 Incidencia_Medida";
                                    //////            workSheet.Cells[f, PosicionSubEstadoGlobal + 1].Value = "02_INCIDENCIA_MEDIDA";
                                    //////            workSheet.Cells[f, PosicionSubEstadoGlobal + 2].Value = "02_INCIDENCIA";
                                    //////            workSheet.Cells[f, PosicionReincidente].Value = reincidente;
                                    //////            ExisteIncidenciaBTN = true;
                                    //////        }
                                           
                                    //////    }
                                      
                                    //////}
                                }

                                //Calculo DIAS ESTADO GLOBAL *************************
                                if (blnLanzarDiasKEE == true) // Sólo se actualizan los datos en la tabla la primera vez, por si hay ejecuciones sucesivas
                                {
                                    if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                    {
                                        DiasExistenciaKEEGlobal = FuncionDiasKEEGlobal(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), workSheet.Cells[f, PosicionSubEstadoKEE].Value.ToString());
                                    }
                                    else
                                    {
                                        DiasExistenciaKEEGlobal = FuncionDiasKEEGlobal(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), workSheet.Cells[f, PosicionSubEstadoSAP].Value.ToString());
                                    }
                                }
                                else
                                {
                                    if (workSheet.Cells[f, PosicionSubEstadoKEE].Value != "")
                                    {
                                        DiasExistenciaKEEGlobal = FuncionDiasKEERecuperarDiasTablaGlobal(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), workSheet.Cells[f, PosicionSubEstadoKEE].Value.ToString());
                                    }
                                    else
                                    {
                                        DiasExistenciaKEEGlobal = FuncionDiasKEERecuperarDiasTablaGlobal(r["cups20"].ToString(), Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd"), Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd"), workSheet.Cells[f, PosicionSubEstadoSAP].Value.ToString());
                                    }
                                }
                                //FIN DIAS ESTADO GLOBAL*******************************


                                // Pego Nº DIAS ESTADO GLOBAL, va a continuación de DIAS ESTADO KEE
                                workSheet.Cells[f, PosicionDiasEstadoKEE + 1].Value = DiasExistenciaKEEGlobal; 

                                //////if (ExisteIncidenciaBTN == true)
                                //////{
                                //////    //No miro el area, ya tengo asignados los estados
                                //////    Auxiliar = "";

                                //////}
                                //////else
                                //////{

                                Auxiliar = "";
                                Auxiliar = EstadosAreas(workSheet.Cells[f, PosicionSubEstadoGlobal].Value.ToString(), cups20, aniomes.ToString(), 3, NumeroDiasKEE, ExisteIncidenciaBTN);

                                if (Auxiliar != "")
                                    {
                                        string[] CadenaEstados = Auxiliar.Split(';');
                                        workSheet.Cells[f, PosicionSubEstadoGlobal +1].Value = CadenaEstados[0];
                                        workSheet.Cells[f, PosicionSubEstadoGlobal + 2].Value = CadenaEstados[1];
                                        workSheet.Cells[f, PosicionSubEstadoGlobal + 3].Value = CadenaEstados[2];
                                    }
                                else
                                    {
                                        workSheet.Cells[f, PosicionSubEstadoGlobal +1].Value = ""; c++;
                                        workSheet.Cells[f, PosicionSubEstadoGlobal + 2].Value = ""; c++;
                                        workSheet.Cells[f, PosicionSubEstadoGlobal + 3].Value = ""; c++;

                                        if (SubestadosNoEncontrados == "")
                                        {
                                            SubestadosNoEncontrados = "Hoja BTN: " + cups20 + " - " + Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd") + " - " + Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd") + " - " + r["de_estado"].ToString() + " -" + r["de_subestado"].ToString() + "- No existe el subestado en la tabla paramétrica";

                                        }
                                        else
                                        {
                                            SubestadosNoEncontrados = SubestadosNoEncontrados + ";Hoja BTN: " + cups20 + " - " + Convert.ToDateTime(r["fh_desde"]).Date.ToString("yyyy-MM-dd") + " - " + Convert.ToDateTime(r["fh_hasta"]).Date.ToString("yyyy-MM-dd") + " - " + r["de_estado"].ToString() + " -" + r["de_subestado"].ToString() + "- No existe el subestado en la tabla paramétrica";
                                        }
                                }
                            } //Fin if (Segmentos.IndexOf(r["cd_segmento_ptg"].ToString()) > 0)
                        } // fin  if (r["cd_segmento_ptg"] != System.DBNull.Value)

                    } // Fin  while (r.Read())
                    db.CloseConnection();

                    intConteoBTNDetalle = f-1;

                    headerCells = workSheet.Cells[1, 1, 1, c];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;

                    //08/02/2024 Paco, mover columna de fecha baja de Kronos antes de fecha de baja de SAP
                    workSheet.InsertColumn(PosicionBajaSap, 1); //Inserta en la posicion 12 una columna en blanco(detrás de FH_BAJA_SAP)
                    workSheet.Cells[1, PosicionBajaKEE, f, PosicionBajaKEE].Copy(workSheet.Cells[1, PosicionBajaSap, f, PosicionBajaSap]); //Copia la columna 31 (Fecha BAJA KEE) en la posición creada anteriormente
                    workSheet.DeleteColumn(PosicionBajaKEE); //Borra la columna 31 Fecha BAJA KEE

                    //Pegamos Subestadoglobal, estado_global y estado global a reportar despues de FH_HASTA *** (!!!OJO, he metido ya por medio FECHA BAJA KEE (Sumo una más)!!)****
                    workSheet.InsertColumn(PosicionPegarSubEstadoGlobal, 1); //Inserta en la posicion 5 una columna en blanco(detrás de FH_HASTA)
                    workSheet.Cells[1, PosicionSubEstadoGlobal + 1, f, PosicionSubEstadoGlobal + 1].Copy(workSheet.Cells[1, PosicionPegarSubEstadoGlobal, f, PosicionPegarSubEstadoGlobal]);

                    workSheet.InsertColumn(PosicionPegarSubEstadoGlobal + 1, 1); //inserto en posición 6 (ya he pegado subestadoglobal)
                    workSheet.Cells[1, PosicionSubEstadoGlobal + 3, f, PosicionSubEstadoGlobal + 3].Copy(workSheet.Cells[1, PosicionPegarSubEstadoGlobal + 1, f, PosicionPegarSubEstadoGlobal + 1]);

                    workSheet.InsertColumn(PosicionPegarSubEstadoGlobal + 2, 1); //inserto en posición 7 (ya he pegado subestadoglobal)
                    workSheet.Cells[1, PosicionSubEstadoGlobal + 5, f, PosicionSubEstadoGlobal + 5].Copy(workSheet.Cells[1, PosicionPegarSubEstadoGlobal + 2, f, PosicionPegarSubEstadoGlobal + 2]);

                    ///borro las originales (tres a la derecha, he insertado tres columnas)
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal + 3);
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal + 3);
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal + 3);

                    //La columna area la pongo detrás de fh_hasta
                    workSheet.InsertColumn(PosicionPegarSubEstadoGlobal, 1);
                    // Ya he quitado lo que originalmente estaba detras de posiciónSubEstado Global, sólo queda el area en esa poscion
                    workSheet.Cells[1, PosicionSubEstadoGlobal + 4, f, PosicionSubEstadoGlobal + 4].Copy(workSheet.Cells[1, PosicionPegarSubEstadoGlobal, f, PosicionPegarSubEstadoGlobal]);
                    workSheet.DeleteColumn(PosicionSubEstadoGlobal + 4);


                    // He quitado  lo que hay entre discrepancias y multipuntos, las posiciones me valen, ahora he metido los globales por delante
                    //////workSheet.InsertColumn(PosicionBajaKEE + 3, 1);
                    //////workSheet.Cells[1, PosicionSubEstadoGlobal + 4, f, PosicionSubEstadoGlobal + 4].Copy(workSheet.Cells[1, PosicionBajaKEE + 3, f, PosicionBajaKEE + 3]); //Copia la columna 31 (Fecha BAJA KEE) en la posición creada anteriormente
                    //////workSheet.DeleteColumn(PosicionSubEstadoGlobal + 4);

                    //workSheet.Cells[1, 5, 100, 5].Copy(workSheet.Cells[1, 2, 100, 2]);
                    //Copia la columna 5 en la columna 2 Básicamente Source.Copy(Destino) Esto solo copiaría las primeras 100 filas.


                    ////////El final es discrepancias seguido de multipuntos, los tengo que cambiar de orden --> discrepancia=48 y multipunto=49 después de las inserciones de columna
                    ////////no lo puedo hacer durante el proceso y poner multipuntos antes, ya que las discrepancias es un proceso aparte y separado
                    ////////Busco donde ha quedado discrepancias
                    //////intColumna = 1;
                    //////while (workSheet.Cells[1, intColumna].Value != "" & workSheet.Cells[1, intColumna].Value != "Discrepancias")
                    //////{
                    //////    intColumna = intColumna + 1;
                    //////}
                    //////workSheet.InsertColumn(intColumna, 1); //Inserta  columna en blanco delante de discrepancias
                    //////workSheet.Cells[1, intColumna + 2, f, intColumna + 2].Copy(workSheet.Cells[1, intColumna, f, intColumna]); //copio multipunto en la columna que he insertado
                    //////workSheet.DeleteColumn(intColumna + 2); //Borra la columna original de multipunto

                    ///////
                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();
                    #endregion

                   
                    if (blnLanzarDiasKEE == true)  //Cierro kee_Dias_Global
                    {
                        ///todos los que están activos y  no se ha marcado PROCESADO_DIARIO = 'S', son puntos que se cierran, actualizo la FH_CIERRE
                        strSql = "update kee_Dias_Global  set"
                                   + " FH_CIERRE = '" + DateTime.Now.ToString("yyyy-MM-dd") + "',"
                                   + " FH_ACTUALIZACION = '" + DateTime.Now.ToString("yyyy-MM-dd") + "',"
                                   + " ACTIVO= False"
                                   + " WHERE  Activo = true"
                                   + " AND FH_ACTUALIZACION < '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";

                        db = new MySQLDB(MySQLDB.Esquemas.MED);
                        command = new MySqlCommand(strSql, db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                    }

                    //Paco 12/02/2024  Guardamos la hoja de "DETALLE ES" y la hoja "Detalle POR MT-BTE"  en una tabla de MySql
                    //Creamos dos nuevas hojas donde mostramos este histórico "RESUMEN ES" y "RESUMEN POR MT-BTE" 
                    int ContadorReplace;
                    StringBuilder sb = new StringBuilder();
                    int intFila = 1;
                    //////int intColumna = 1;
                    string Fecha;
                    int Resumen;
                    string NombreHojaResumen;
                    string nombreColumna;
                    int CuentaColumna;
                    NombreHojaResumen = "";

                    ///Paco 04-04-2025 Grabamos los datos de las hojas de detalle en tabla facturacionb2b_owner.sap_kee_resumen_es_por del Esquema de facturación
                    ////// Primero un delete por si se lanza más  de una vez
                    strSqlEsqFact = "DELETE from  facturacionb2b_owner.sap_kee_resumen_es_por where FH_CARGA='" + DateTime.Now.ToString("yyyy-MM-dd") + "'";
                    dbRS = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD_FACTURACION);
                    commandRS = new OdbcCommand(strSqlEsqFact, dbRS.con);
                    rRS = commandRS.ExecuteReader();
                    dbRS.CloseConnection();
                    strSqlEsqFact = "";
                    // Fin Paco 04-04-2025


                    //Guardamos las hojas de detalle en la tabla sap_KEE_Resumen_ES_POR_MT_BTE_New y facturacionb2b_owner.sap_kee_resumen_es_por
                    for (Resumen = 1; Resumen < 4; Resumen++){

                        intFila = 1;
                        intColumna = 1;

                        if (Resumen == 1) {
                            NombreHojaResumen = "Detalle ES";

                        }
                        if (Resumen == 2)
                        {
                            NombreHojaResumen = "Detalle POR MT-BTE";
                        }
                        if (Resumen == 3)
                        {
                            NombreHojaResumen = "Detalle POR BTN";
                        }
                       
                        var hojaActual = excelPackage.Workbook.Worksheets[NombreHojaResumen];
                        while (hojaActual.Cells[1, intColumna].Value != "" & hojaActual.Cells[1, intColumna].Value != null)
                        {
                            intColumna = intColumna + 1;
                        }
                        intColumna = intColumna - 1;
                        //31
                        while (hojaActual.Cells[intFila, 1].Value != "" & hojaActual.Cells[intFila, 1].Value != null & hojaActual.Cells[intFila, 2].Value != "" & hojaActual.Cells[intFila, 2].Value != null)
                        {
                            intFila = intFila + 1;
                        }
                        intFila = intFila - 1;

                        //OJO!! Esto funciona para columna con más de una letra
                        if (hojaActual.Cells[1, intColumna].Address.Length > 2 )
                            nombreColumna = hojaActual.Cells[1, intColumna].Address.Substring(0, 2);
                        else
                            nombreColumna = hojaActual.Cells[1, intColumna].Address.Substring(0, 1);

                        //////if (Resumen == 1)
                        //////{
                        //////    nombreColumna = String.Concat("A2:AQ", intFila.ToString());
                        //////}
                        //////else {
                        //////    if (Resumen == 2)
                        //////        nombreColumna = String.Concat("A2:AR", intFila.ToString());
                        //////    else
                        //////        nombreColumna = String.Concat("A2:AF", intFila.ToString());
                        //////}
                        nombreColumna = String.Concat("A2:" , nombreColumna, intFila.ToString());
                        var rango= hojaActual.Cells[nombreColumna];

                        ContadorReplace = 0;
                        for (
                            int i = 2; i <= intFila; i++)
                        {
                            CuentaColumna = 1;
                            if (ContadorReplace == 0)
                            {

                                if (Resumen == 1 || Resumen == 2)
                                {
                                    if (Resumen == 1)
                                    {
                                        sb.Append("REPLACE INTO sap_KEE_Resumen_ES_POR_MT_BTE_New( CUPS20,PERIODO, mes,FH_DESDE,FH_HASTA, AREA_ESTADO, Subestado_global, ESTADO_GLOBAL, ESTADO_GLOBAL_A_REPORTAR,IMPORTE_PDTE_FACTURAR,MESES_PDTES_FACTURAR,Agora,cliente, DIAS_ESTADO_SAP, DIAS_ESTADO_KEE, DIAS_ESTADO_GLOBAL, Incidencia_Facturacion, Estado_FAC_SE, Titulo_FAC, Incidencia_Medida, Reincidente, Estado_incidencia, Fecha_Apertura, Prioridad, Titulo, E_S_Estado,FH_ALTA_SALESFORCE, Fecha_Alta_KEE, FH_ALTA_SAP, FH_BAJA_SALESFORCE,Fecha_BAJA_KEE,FH_BAJA_SAP"
                                         + ",Empresa, tipo_cliente,nif,FPSERCON,N_INSTALACION,tarifa,contrato,Distribuidora,Estado,Subestado,Tam,ULT_FH_DESDE_FACTURADA, ULT_FH_HASTA_FACTURADA,Estado_periodo_KEE,Area_responsable_KEE,Subestado_KEE,Estado_KEE,FH_DESDE_KEE,FH_HASTA_KEE,Multipunto,Discrepancia, FH_CARGA) values ('");
                                    }

                                    else
                                    {
                                        sb.Append("REPLACE INTO sap_KEE_Resumen_ES_POR_MT_BTE_New( CUPS20,PERIODO, mes,FH_DESDE,FH_HASTA, AREA_ESTADO, Subestado_global, ESTADO_GLOBAL, ESTADO_GLOBAL_A_REPORTAR,IMPORTE_PDTE_FACTURAR,MESES_PDTES_FACTURAR,Agora,cliente, DIAS_ESTADO_SAP, DIAS_ESTADO_KEE, DIAS_ESTADO_GLOBAL, Incidencia_Facturacion, Estado_FAC_SE, Titulo_FAC,Incidencia_Medida, Reincidente, Estado_incidencia, Fecha_Apertura, Prioridad, Titulo, E_S_Estado,FH_ALTA_SALESFORCE,Fecha_Alta_KEE,FH_ALTA_SAP, FH_BAJA_SALESFORCE,Fecha_BAJA_KEE,FH_BAJA_SAP"
                                        + ",Empresa,segmento, tipo_cliente,nif,FPSERCON,N_INSTALACION,tarifa,contrato,Distribuidora,Estado,Subestado,Tam,ULT_FH_DESDE_FACTURADA, ULT_FH_HASTA_FACTURADA,Estado_periodo_KEE,Area_responsable_KEE,Subestado_KEE,Estado_KEE,FH_DESDE_KEE,FH_HASTA_KEE,Multipunto, Discrepancia,FH_CARGA) values ('");
                                    }
                                }
                                else {
                                  //////  sb.Append("REPLACE INTO sap_KEE_Resumen_ES_POR_MT_BTE(Empresa,segmento, tipo_cliente,nif,cliente,FALTACONT,FPSERCON,CUPS20,FH_DESDE,FH_HASTA,FH_BAJA_SALESFORCE,FH_BAJA_SAP,Fecha_BAJA_KEE,"
                                  //////+ " N_INSTALACION,tarifa,contrato,mes,Distribuidora,Estado,Subestado,DIAS_ESTADO,Tam,MESES_PDTES_FACTURAR,ULT_FH_DESDE_FACTURADA,"
                                  //////+ " ULT_FH_HASTA_FACTURADA,IMPORTE_PDTE_FACTURAR,Agora,Estado_periodo_KEE,Area_responsable_KEE,Subestado_KEE,Estado_KEE,FH_DESDE_KEE,FH_HASTA_KEE,Discrepancia, Subestado_global, Incidencia, Estado_incidencia, Fecha_Apertura, Prioridad, Titulo, E_S_Estado, ESTADO_GLOBAL, ESTADO_GLOBAL_A_REPORTAR, Multipunto,FH_CARGA) values ('");

                                    sb.Append("REPLACE INTO sap_KEE_Resumen_ES_POR_MT_BTE_New( CUPS20,PERIODO, mes,FH_DESDE,FH_HASTA, AREA_ESTADO,Subestado_global, ESTADO_GLOBAL, ESTADO_GLOBAL_A_REPORTAR,IMPORTE_PDTE_FACTURAR,MESES_PDTES_FACTURAR,Agora,cliente, DIAS_ESTADO_SAP, DIAS_ESTADO_KEE, DIAS_ESTADO_GLOBAL, Incidencia_Facturacion, Estado_FAC_SE, Titulo_FAC,Incidencia_Medida, Reincidente, Estado_incidencia, Fecha_Apertura, Prioridad, Titulo, E_S_Estado,FH_ALTA_SALESFORCE,Fecha_Alta_KEE,FH_ALTA_SAP, FH_BAJA_SALESFORCE,Fecha_BAJA_KEE,FH_BAJA_SAP"
                                  + ",Empresa,segmento, tipo_cliente,nif,FPSERCON,N_INSTALACION,tarifa,contrato,Distribuidora,Estado,Subestado,Tam,ULT_FH_DESDE_FACTURADA, ULT_FH_HASTA_FACTURADA,Estado_periodo_KEE,Area_responsable_KEE,Subestado_KEE,Estado_KEE,FH_DESDE_KEE,FH_HASTA_KEE, Discrepancia, NDIAS_PTE_KEE,FH_CARGA) values ('");
                                }

                                //////if (Resumen == 1 || Resumen == 2)
                                //////{
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //CUPS20
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //PERIODO
                                    sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //MES
                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_DESDE
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                    else
                                        sb.Append("null,");
                                    CuentaColumna++;
                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_HASTA
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("','");
                                    else
                                        sb.Append("null,'");
                                    CuentaColumna++;
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //AREA_ESTADO
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Subestado global
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //ESTADO GLOBAL
                                    sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //ESTADO GLOBAL A REPORTAR
                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //IMPORTE PDTE FACTURAR
                                        sb.Append(Convert.ToString(rango[i, CuentaColumna].Value).Replace(",", ".")).Append(",");
                                    else
                                        sb.Append("null,");
                                    CuentaColumna++;
                                    sb.Append(rango[i, CuentaColumna].Value).Append(",'"); CuentaColumna++; //MESES PDTES FACTURAR
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //AGORA
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //CLIENTE

                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //DIAS_ESTADO_SAP
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //DIAS_ESTADO_KEE
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //DIAS_ESTADO_GLOBAL
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Incidencia_Facturacion
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Estado_FAC_SE
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Titulo_FAC
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Incidencia_Medida
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Reincidente
                                    sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //Estado Incidencia
                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_APERTURA
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("','");
                                    else
                                        sb.Append("null,'");
                                    CuentaColumna++;

                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Prioridad
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Titulo
                                    sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //ES_Estado

                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //Fecha_ALTA_SALESFORCE
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                    else
                                        sb.Append("null,");
                                    CuentaColumna++;

                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //Fecha_ALTA_KEE
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                    else
                                        sb.Append("null,");
                                    CuentaColumna++;

                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_ALTA_SAP
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                    else
                                        sb.Append("null,");
                                    CuentaColumna++;
                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_BAJA_SALESFORCE
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                    else
                                        sb.Append("null,");
                                    CuentaColumna++;
                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_BAJA KEE
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                    else
                                        sb.Append("null,");
                                    CuentaColumna++;
                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_BAJA SAP
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("','");
                                    else
                                        sb.Append("null,'");
                                    CuentaColumna++;

                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Empresa
                                    if (Resumen != 1) {//SEGMENTO   
                                        //////sb.Append("','");   
                                        sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++;
                                    }

                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Tipo Cliente
                                    sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //NIF

                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FPSERCON
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("','");
                                    else
                                        sb.Append("null,'");
                                    CuentaColumna++;
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Nº INSTALACION
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //TARIFA
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //contrato
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //distribuidora
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Estado
                                    sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //SubEstado
                                    

                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //TAM
                                        sb.Append(Convert.ToString(rango[i, CuentaColumna].Value).Replace(",", ".")).Append(",");
                                    else
                                        sb.Append("null,");
                                    CuentaColumna++;

                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //ULT FH DESDE FACTURADA
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                    else
                                        sb.Append("null,");
                                    CuentaColumna++;
                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //ULT FH HASTA FACTURADA
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("','");
                                    else
                                        sb.Append("null,'");
                                    CuentaColumna++;
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Estado Periodo KEE
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Area Responsable KEE
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Subestado KEE
                                    sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //Estado KEE
                                    //FH_DESDE_KEE y FH_HASTA_KEE
                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "")
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                    else
                                        sb.Append("null,");
                                    CuentaColumna++;

                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "")
                                        sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("','");
                                    else
                                        sb.Append("null,'");
                                    CuentaColumna++;


                                    if (Resumen == 1 || Resumen == 2)
                                    {
                                        sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Multipunto
                                    }

                                    sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //Discrepancias
                                                                                                            


                                    if (Resumen == 3)
                                    //NDIAS_ESTADO_KEE y NDIAS_PTE_KEE
                              
                                    {
                                    //////if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "")
                                    //////    sb.Append(rango[i, CuentaColumna].Value).Append(",");
                                    //////else
                                    //////    sb.Append("null,");

                                    //////CuentaColumna++;

                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "")
                                            sb.Append(rango[i, CuentaColumna].Value).Append(",");
                                        else
                                            sb.Append("null,");
                                        CuentaColumna++;
                                    }
                                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd")).Append("')");

                            }
                            else
                            {


                                sb.Append(",('");
                                //////if (Resumen == 1 || Resumen == 2)
                                //////{
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //CUPS20
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //PERIODO
                                sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //MES
                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_DESDE
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                else
                                    sb.Append("null,");
                                CuentaColumna++;
                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_HASTA
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("','");
                                else
                                    sb.Append("null,'");
                                CuentaColumna++;
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //AREA_ESTADO
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Subestado global
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //ESTADO GLOBAL
                                sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //ESTADO GLOBAL A REPORTAR
                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //IMPORTE PDTE FACTURAR
                                    sb.Append(Convert.ToString(rango[i, CuentaColumna].Value).Replace(",", ".")).Append(",");
                                else
                                    sb.Append("null,");
                                CuentaColumna++;
                                sb.Append(rango[i, CuentaColumna].Value).Append(",'"); CuentaColumna++; //MESES PDTES FACTURAR
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //AGORA
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //CLIENTE

                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //DIAS_ESTADO_SAP
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //DIAS_ESTADO_KEE
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //DIAS_ESTADO_GLOBAL
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Incidencia_Facturacion
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Estado_FAC_SE
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Titulo_FAC
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Incidencia_Medida
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Reincidente
                                sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //Estado Incidencia
                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_APERTURA
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("','");
                                else
                                    sb.Append("null,'");
                                CuentaColumna++;

                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Prioridad
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Titulo
                                sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //ES_Estado

                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //Fecha_ALTA_SALESFORCE
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                else
                                    sb.Append("null,");
                                CuentaColumna++;

                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //Fecha_ALTA_KEE
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                else
                                    sb.Append("null,");
                                CuentaColumna++;

                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_ALTA_SAP
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                else
                                    sb.Append("null,");
                                CuentaColumna++;
                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_BAJA_SALESFORCE
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                else
                                    sb.Append("null,");
                                CuentaColumna++;
                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_BAJA KEE
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                else
                                    sb.Append("null,");
                                CuentaColumna++;
                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FH_BAJA SAP
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("','");
                                else
                                    sb.Append("null,'");
                                CuentaColumna++;

                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Empresa
                                if (Resumen != 1)
                                {//SEGMENTO   
                                 //////sb.Append("','");   
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++;
                                }

                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Tipo Cliente
                                sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //NIF

                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //FPSERCON
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("','");
                                else
                                    sb.Append("null,'");
                                CuentaColumna++;
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Nº INSTALACION
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //TARIFA
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //contrato
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //distribuidora
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Estado
                                sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //SubEstado
                                //////sb.Append(rango[i, CuentaColumna].Value).Append(","); CuentaColumna++; //Dias_Estado

                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //TAM
                                    sb.Append(Convert.ToString(rango[i, CuentaColumna].Value).Replace(",", ".")).Append(",");
                                else
                                    sb.Append("null,");
                                CuentaColumna++;

                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //ULT FH DESDE FACTURADA
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                else
                                    sb.Append("null,");
                                CuentaColumna++;
                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "") //ULT FH HASTA FACTURADA
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("','");
                                else
                                    sb.Append("null,'");
                                CuentaColumna++;
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Estado Periodo KEE
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Area Responsable KEE
                                sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Subestado KEE
                                sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //Estado KEE
                                                                                                        //FH_DESDE_KEE y FH_HASTA_KEE
                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "")
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("',");
                                else
                                    sb.Append("null,");
                                CuentaColumna++;

                                if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "")
                                    sb.Append("'").Append(Convert.ToDateTime(rango[i, CuentaColumna].Value).ToString("yyyy-MM-dd")).Append("','");
                                else
                                    sb.Append("null,'");
                                CuentaColumna++;

                                if (Resumen == 1 || Resumen == 2)
                                {
                                    sb.Append(rango[i, CuentaColumna].Value).Append("','"); CuentaColumna++; //Multipunto
                                }

                                sb.Append(rango[i, CuentaColumna].Value).Append("',"); CuentaColumna++; //Discrepancias


                                if (Resumen == 3)
                                //NDIAS_ESTADO_KEE y NDIAS_PTE_KEE

                                {
                                    //////if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "")
                                    //////    sb.Append(rango[i, CuentaColumna].Value).Append(",");
                                    //////else
                                    //////    sb.Append("null,");

                                    //////CuentaColumna++;

                                    if (rango[i, CuentaColumna].Value != null & rango[i, CuentaColumna].Value != "")
                                        sb.Append(rango[i, CuentaColumna].Value).Append(",");
                                    else
                                        sb.Append("null,");
                                    CuentaColumna++;
                                }
                                sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd")).Append("')");

                            }
                            ContadorReplace = ContadorReplace + 1;
                            if (ContadorReplace == 100)
                            {
                                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                                //////Console.WriteLine(sb.ToString());
                                command = new MySqlCommand(sb.ToString(), db.con);
                                command.ExecuteNonQuery();

                                ///Paco 04-04-2025 Grabamos los datos de las hojas de detalle en tabla facturacionb2b_owner.sap_kee_resumen_es_por del Esquema de facturación
                                strSqlEsqFact = sb.ToString();
                                strSqlEsqFact = strSqlEsqFact.Replace("REPLACE INTO sap_KEE_Resumen_ES_POR_MT_BTE_New", "INSERT INTO facturacionb2b_owner.sap_kee_resumen_es_por");
                                dbRS = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD_FACTURACION);
                                commandRS = new OdbcCommand(strSqlEsqFact, dbRS.con);
                                rRS = commandRS.ExecuteReader();
                                dbRS.CloseConnection();
                                strSqlEsqFact = "";
                                /// Fin Paco 04-04-2025

                                sb = null;
                                sb = new StringBuilder();
                                ContadorReplace = 0;
                                db.CloseConnection();
                            }
                        }

                        if (ContadorReplace >0)
                        {
                            db = new MySQLDB(MySQLDB.Esquemas.FAC);
                            command = new MySqlCommand(sb.ToString(), db.con);
                            command.ExecuteNonQuery();

                            ///Paco 04-04-2025 Grabamos los datos de las hojas de detalle en tabla facturacionb2b_owner.sap_kee_resumen_es_por del Esquema de facturación
                            strSqlEsqFact = sb.ToString();
                            strSqlEsqFact = strSqlEsqFact.Replace("REPLACE INTO sap_KEE_Resumen_ES_POR_MT_BTE_New", "INSERT INTO facturacionb2b_owner.sap_kee_resumen_es_por");
                            dbRS = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD_FACTURACION);
                            commandRS = new OdbcCommand(strSqlEsqFact, dbRS.con);
                            rRS = commandRS.ExecuteReader();
                            dbRS.CloseConnection();
                            strSqlEsqFact = "";
                            /// Fin Paco 04-04-2025

                            sb = null;
                            sb = new StringBuilder();
                            ContadorReplace = 0;
                            db.CloseConnection();
                        }

                    }

                    ///29/08/2024  COLUMNAS (CODIGO NUEVO PARA ACELERAR EXCEL, más adelante usamos  Aspose.Cells  para luego separar por ; la columna A que tiene todos los campos
                    TotalPruebaES = 0;
                    TotalPruebaMTBTE = 0;
                    TotalPruebaBTN = 0;

                    //28/04/2025 CREAMOS HOJAS/CUADROS RESUMEN a PARTIR DE TABLA  facturacionb2b_owner.sap_kee_resumen_es_por 
                    for (Hoja = 1; Hoja < 4; Hoja++)
                    {
                        PintoDias = true;
                        TotalAgoraDia1 = 0;
                        TotalAgoraDia2 = 0;
                        TotalAgoraDia3 = 0;
                        TotalAgoraDia4 = 0;
                        TotalAgoraDia5 = 0;
                        TotalImporteDia1 = 0;
                        TotalImporteDia2 = 0;
                        TotalImporteDia3 = 0;
                        TotalImporteDia4 = 0;
                        TotalImporteDia5 = 0;
                        TotalAgoraDia1Final = 0;
                        TotalAgoraDia2Final = 0;
                        TotalAgoraDia3Final = 0;
                        TotalAgoraDia4Final = 0;
                        TotalAgoraDia5Final = 0;
                        TotalImporteDia1Final = 0;
                        TotalImporteDia2Final = 0;
                        TotalImporteDia3Final = 0;
                        TotalImporteDia4Final = 0;
                        TotalImporteDia5Final = 0;
                        int InicioCaja = 0;

                        if (Hoja == 1)
                        {
                            workSheet = excelPackage.Workbook.Worksheets.Add("Resumen ES");
                            strSqlHoja = "select estado, subestado_global, agora, count(agora) as Conteo, sum(importe_pdte_Facturar) as Importe, fh_carga"
                            + " from facturacionb2b_owner.sap_kee_resumen_es_por"
                            + " where (segmento='' or segmento is null)"
                            + " and fh_carga >= "
                            + " ("
                            + "   SELECT MIN(A.FH_CARGA) FROM"
                            + "   ("
                            + "     SELECT distinct FH_CARGA"
                            + "     FROM facturacionb2b_owner.sap_kee_resumen_es_por"
                            + "     ORDER BY FH_CARGA DESC"
                            + "     LIMIT 5"
                            + "   ) AS A"
                            + " )"
                            + " group by estado,subestado_global, agora, fh_carga"
                            + " order by  agora asc,estado asc, subestado_global asc, fh_carga asc";
                        }
                        if (Hoja == 2)
                        {
                            workSheet = excelPackage.Workbook.Worksheets.Add("Resumen POR MT-BTE");
                            strSqlHoja = "select estado, subestado_global, agora, count(agora) as Conteo, sum(importe_pdte_Facturar) as Importe, fh_carga"
                            + " from facturacionb2b_owner.sap_kee_resumen_es_por"
                            + " where segmento in ('MT','BTE')"
                            + " and fh_carga >= "
                            + " ("
                            + "   SELECT MIN(A.FH_CARGA) FROM"
                            + "   ("
                            + "     SELECT distinct FH_CARGA"
                            + "     FROM facturacionb2b_owner.sap_kee_resumen_es_por"
                            + "     ORDER BY FH_CARGA DESC"
                            + "     LIMIT 5"
                            + "   ) AS A"
                            + " )"
                            + " group by estado,subestado_global, agora, fh_carga"
                            + " order by  agora asc,estado asc, subestado_global asc, fh_carga asc";                     
                        }
                        if (Hoja == 3)
                        {
                            workSheet = excelPackage.Workbook.Worksheets.Add("Resumen POR BTN");
                            strSqlHoja = "select estado, subestado_global, agora, count(agora) as Conteo, sum(importe_pdte_Facturar) as Importe, fh_carga"
                            + " from facturacionb2b_owner.sap_kee_resumen_es_por"
                            + " where segmento ='BTN'"
                            + " and fh_carga >= "
                            + " ("
                            + "   SELECT MIN(A.FH_CARGA) FROM"
                            + "   ("
                            + "     SELECT distinct FH_CARGA"
                            + "     FROM facturacionb2b_owner.sap_kee_resumen_es_por"
                            + "     ORDER BY FH_CARGA DESC"
                            + "     LIMIT 5"
                            + "   ) AS A"
                            + " )"
                            + " group by estado,subestado_global, agora, fh_carga"
                            + " order by  agora asc,estado asc, subestado_global asc, fh_carga asc";
                        }
                        headerCells = workSheet.Cells[1, 1, 1, 17];
                        headerFont = headerCells.Style.Font;
                        f = 1;
                        dia = 0;
                        //Pego etiquetas
                        workSheet.Cells[1, 1].Value = "INFORME SEGUIMIENTO PENDIENTE FACTURACION TOTAL";
                        workSheet.Cells[1, 1].Style.Font.Bold = true;
                        workSheet.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor(colorDeCeldaTitulo);
                        workSheet.Cells["A1:D1"].Merge = true;
                        workSheet.Cells["A1:D1"].Style.WrapText = true;

                        workSheet.Cells[3, 1].Value = "ÁGORA (SÍ/NO)";
                        workSheet.Cells[3, 1].Style.Font.Bold = true;
                        //////workSheet.Column(1).Width = 50;
                        //////workSheet.Cells["A3"].AutoFitColumns(200);
                        workSheet.Cells["A3:A4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells["A3:A4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells["A3:A4"].Style.Fill.BackgroundColor.SetColor(colorDeCeldaCabecera);
                        workSheet.Cells["A3:A4"].Merge = true;
                        workSheet.Cells["A3:A4"].Style.WrapText = true;
                        workSheet.Cells["A3:A4"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        workSheet.Cells[3, 2].Value = "RESPONSABLE";
                        workSheet.Cells[3, 2].Style.Font.Bold = true;
                        
                        //////workSheet.Cells["B3"].AutoFitColumns(250);
                        workSheet.Cells["B3:B4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells["B3:B4"].Style.Fill.BackgroundColor.SetColor(colorDeCeldaCabecera);
                        workSheet.Cells["B3:B4"].Merge = true;
                        workSheet.Cells["B3:B4"].Style.WrapText = true;
                        workSheet.Cells["B3:B4"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        workSheet.Cells["B3:B4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        workSheet.Cells[3, 3].Value = "SUBESTADO";
                        workSheet.Cells[3, 3].Style.Font.Bold = true;
                        workSheet.Cells["C3:C4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells["C3:C4"].Style.Fill.BackgroundColor.SetColor(colorDeCeldaCabecera);
                        workSheet.Cells["C3:C4"].Merge = true;
                        workSheet.Cells["C3:C4"].Style.WrapText = true;
                        workSheet.Cells["C3:C4"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        workSheet.Cells["C3:C4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        workSheet.Cells[3, 4].Value = "Pendiente PS";
                        workSheet.Cells[3, 4].Style.Font.Bold = true;
                        workSheet.Cells["D3:H3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells["D3:H3"].Style.Fill.BackgroundColor.SetColor(colorDeCeldaCabecera);
                        workSheet.Cells["D3:H3"].Merge = true;
                        workSheet.Cells["D3:H3"].Style.WrapText = true;
                        workSheet.Cells["D3:H3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        workSheet.Cells["D3:H3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        if (PintoDias)
                        {
                            strSql = " select A.FH_CARGA from (SELECT distinct FH_CARGA"
                            + "     FROM facturacionb2b_owner.sap_kee_resumen_es_por"
                            + "     ORDER BY FH_CARGA desc"
                            + "     LIMIT 5)  as A order by A.FH_CARGA asc";

                            dbRS = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD_FACTURACION);
                            commandRS = new OdbcCommand(strSql, dbRS.con);
                            rRS = commandRS.ExecuteReader();

                            t = 4;
                            while (rRS.Read())
                            {
                                DiaPintar= Convert.ToDateTime(rRS["FH_CARGA"]);
                                workSheet.Cells[4, t].Value = DiaPintar.ToString("dd/MM/yyyy"); 
                                workSheet.Cells[4, t].Style.Font.Bold = true;
                                workSheet.Cells[4, t].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[4, t].Style.Fill.BackgroundColor.SetColor(colorDeCeldaCabecera);
                                workSheet.Cells[4, t].Merge = true;
                                workSheet.Cells[4, t].Style.WrapText = true;
                                workSheet.Cells[4, t].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                workSheet.Cells[4, t].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                workSheet.Cells[4, t + 5].Value = DiaPintar.ToString("dd/MM/yyyy");
                                workSheet.Cells[4, t + 5].Style.Font.Bold = true;
                                workSheet.Cells[4, t + 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[4, t + 5].Style.Fill.BackgroundColor.SetColor(colorDeCeldaCabecera);
                                workSheet.Cells[4, t + 5].Merge = true;
                                workSheet.Cells[4, t + 5].Style.WrapText = true;
                                workSheet.Cells[4, t + 5].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                workSheet.Cells[4, t + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                //Almaceno los días
                                switch (t)
                                {
                                    case 4:
                                        Dia1Date = Convert.ToDateTime(rRS["FH_CARGA"]);
                                        Dia1= Dia1Date.ToString("dd/MM/yyyy");
                                        break;
                                    case 5:
                                        Dia2Date = Convert.ToDateTime(rRS["FH_CARGA"]);
                                        Dia2 = Dia2Date.ToString("dd/MM/yyyy");
                                        break;
                                    case 6:
                                        Dia3Date = Convert.ToDateTime(rRS["FH_CARGA"]);
                                        Dia3 = Dia3Date.ToString("dd/MM/yyyy");
                                        break;
                                    case 7:
                                        Dia4Date = Convert.ToDateTime(rRS["FH_CARGA"]);
                                        Dia4 = Dia4Date.ToString("dd/MM/yyyy");
                                        break;
                                    case 8:
                                        Dia5Date = Convert.ToDateTime(rRS["FH_CARGA"]);
                                        Dia5 = Dia5Date.ToString("dd/MM/yyyy");
                                        break;
                                }
                                t = t + 1;
                            }
                            dbRS.CloseConnection();
                            PintoDias = false;
                        }

                        workSheet.Cells[3, 9].Value = "Pendiente Económico (€)";
                        workSheet.Cells[3, 9].Style.Font.Bold = true;
                        workSheet.Cells["I3:M3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells["I3:M3"].Style.Fill.BackgroundColor.SetColor(colorDeCeldaCabecera);
                        workSheet.Cells["I3:M3"].Merge = true;
                        workSheet.Cells["I3:M3"].Style.WrapText = true;
                        workSheet.Cells["I3:M3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        workSheet.Cells["I3:M3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                       
                        dbRS = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD_FACTURACION);
                        commandRS = new OdbcCommand(strSqlHoja, dbRS.con);
                        rRS = commandRS.ExecuteReader();

                        f = 5;

                        ////allCells = workSheet.Cells[1, 1, f, 13];
                        ////allCells.AutoFitColumns();

                        CambioAgora = true;
                        CambioEstado = true;
                        CambioSubEstado = true;
                        Subestado = "";
                        Estado = "";
                        Agora = "";
                        InicioCaja = 5;
                        FilaCambioAgora = 0;

                        while (rRS.Read())
                        {
                            // Testeo Campo Agora
                            if (Agora == "" || Agora == rRS["agora"].ToString())
                            {
                                Agora = rRS["agora"].ToString();
                                if (CambioAgora)
                                {
                                    if (Agora == "N")
                                    {
                                        workSheet.Cells[f, 1].Value = "NO ÁGORA";
                                    }
                                    else {
                                        workSheet.Cells[f, 1].Value = "SI ÁGORA";
                                    }
                                    workSheet.Cells[f, 1].Style.Font.Bold = true;
                                }
                                CambioAgora = false;
                            }
                            else
                            {
                                CambioAgora = true;

                                //Pinto el total del subestado que no he pintado por el cambio
                                ///Ajustamos la caja con el texto
                                RangoInterno = "B" + InicioCaja.ToString() + ":B" + f.ToString();
                                workSheet.Cells[RangoInterno].Merge = true;
                                workSheet.Cells[RangoInterno].Style.WrapText = true;
                                workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);  
                                InicioCaja = f;

                                //Cambio Estado Aumento fila
                                f = f + 1;

                                //Pinto las lineas de totales en gris
                                workSheet.Cells[f, 3].Value = "Total " + Estado;
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                //////RangoPintoGris = "B" + f.ToString() + ":C" + f.ToString();
                                //////workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                RangoPintoGris = "B" + f.ToString() + ":M" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaGris);
                                workSheet.Cells[RangoPintoGris].Style.Font.Bold = true;
                                workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);



                                //////workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                //////workSheet.Cells[f, c].Style.Font.Bold = true;
                                //////workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                //Cambia el estado: Pego la suma de conteo y de importe de cada dia
                                for (b = 4; b < 9; b++)
                                {
                                    switch (b)
                                    {
                                        case 4:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia1;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia1;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b+5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b+5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            TotalAgoraDia1 = 0;
                                            TotalImporteDia1 = 0;
                                            break;
                                        case 5:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia2;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia2;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            TotalAgoraDia2 = 0;
                                            TotalImporteDia2 = 0;
                                            break;
                                        case 6:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia3;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia3;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            TotalAgoraDia3 = 0;
                                            TotalImporteDia3 = 0;
                                            break;
                                        case 7:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia4;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia4;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            TotalAgoraDia4 = 0;
                                            TotalImporteDia4 = 0;
                                            break;
                                        case 8:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia5;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia5;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            TotalAgoraDia5 = 0;
                                            TotalImporteDia5 = 0;
                                            break;
                                    }
                                }

                                //fin Pinto el total del subestado que no he pintado por el cambio de agora


                                //Cambio Agora Aumento fila
                                f = f + 1;
                                //Pinto las lineas de totales No Agora en verde
                                workSheet.Cells[f, 3].Value = "Total No Ágora";
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                //////RangoPintoGris = "A" + f.ToString() + ":C" + f.ToString();
                                //////workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                RangoPintoGris = "A" + f.ToString() + ":M" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaTitulo);
                                workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                workSheet.Cells[RangoPintoGris].Style.Font.Bold = true;

                                FilaCambioAgora = f;
                                ///Ajustamos la caja con el texto
                                RangoInterno = "A5:A" + (FilaCambioAgora-1).ToString();
                                workSheet.Cells[RangoInterno].Merge = true;
                                workSheet.Cells[RangoInterno].Style.WrapText = true;
                                workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                InicioCaja = f + 1;

                                //Cambia el estado: Pego la suma de conteo y de importe de cada dia
                                for (b = 4; b < 9; b++)
                                {
                                    switch (b)
                                    {
                                        case 4:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia1Final;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia1Final;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            break;
                                        case 5:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia2Final;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia2Final;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            break;
                                        case 6:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia3Final;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia3Final;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            break;
                                        case 7:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia4Final;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia4Final;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            break;
                                        case 8:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia5Final;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia5Final;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            break;
                                    }
                                }
                                
                                f = f + 1;
                                Agora = rRS["agora"].ToString();
                                if (CambioAgora)
                                {
                                    if (Agora == "N")
                                    {
                                        workSheet.Cells[f, 1].Value = "NO ÁGORA";
                                    }
                                    else
                                    {
                                        workSheet.Cells[f, 1].Value = "SI ÁGORA";
                                    }
                                    workSheet.Cells[f, 1].Style.Font.Bold = true;
                                }
                                Subestado = "";
                                Estado = "";
                                CambioEstado = true;
                                CambioSubEstado = true;
                                TotalAgoraDia1 = 0;
                                TotalImporteDia1 = 0;
                                TotalAgoraDia2 = 0;
                                TotalImporteDia2 = 0;
                                TotalAgoraDia3 = 0;
                                TotalImporteDia3 = 0;
                                TotalAgoraDia4 = 0;
                                TotalImporteDia4 = 0;
                                TotalAgoraDia5 = 0;
                                TotalImporteDia5 = 0;
                                CambioEstado = true;
                                CambioSubEstado = true;
                            } // Final Testeo Campo Agora

                            //Testeo campo estado
                            if (Estado == "" || Estado == rRS["Estado"].ToString())
                            {
                                Estado = rRS["Estado"].ToString();
                                if (CambioEstado)
                                {
                                    workSheet.Cells[f, 2].Value = Estado;
                                    workSheet.Cells[f, 2].Style.Font.Bold = true;
                                }
                                CambioEstado = false;
                            }
                            else
                            {
                                CambioEstado = true;

                                ///Ajustamos la caja con el texto
                                RangoInterno = "B" + InicioCaja.ToString() + ":B" + f.ToString();
                                workSheet.Cells[RangoInterno].Merge = true;
                                workSheet.Cells[RangoInterno].Style.WrapText = true;
                                workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                InicioCaja = f;

                                //Cambio Estado Aumento fila
                                f = f + 1;
                                InicioCaja = f+1;

                                //Pinto las lineas de totales en gris
                                workSheet.Cells[f, 3].Value = "Total " + Estado;
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                //////RangoPintoGris = "B" + f.ToString() + ":C" + f.ToString();
                                //////workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                RangoPintoGris = "B" + f.ToString() + ":M" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaGris);
                                //////workSheet.Cells[f, c].Value = TotalConteoResponsable;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                //Cambia el estado: Pego la suma de conteo y de importe de cada dia
                                for (b = 4; b < 9;b++)
                                {
                                    switch (b)
                                    {
                                        case 4:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia1;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia1;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            TotalAgoraDia1 = 0;
                                            TotalImporteDia1 = 0;
                                            break;
                                        case 5:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia2;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia2;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            TotalAgoraDia2 = 0;
                                            TotalImporteDia2 = 0;
                                            break;
                                        case 6:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia3;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia3;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            TotalAgoraDia3 = 0;
                                            TotalImporteDia3 = 0;
                                            break;
                                        case 7:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia4;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia4;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            TotalAgoraDia4 = 0;
                                            TotalImporteDia4 = 0;
                                            break;
                                        case 8:
                                            workSheet.Cells[f, b].Value = TotalAgoraDia5;
                                            workSheet.Cells[f, b + 5].Value = TotalImporteDia5;
                                            workSheet.Cells[f, b].Style.Font.Bold = true;
                                            workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                            workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            TotalAgoraDia5 = 0;
                                            TotalImporteDia5 = 0;
                                            break;
                                    }
                                }
                                f = f + 1;
                                Estado = rRS["Estado"].ToString();
                                if (CambioEstado)
                                {
                                    workSheet.Cells[f, 2].Value = Estado;
                                    workSheet.Cells[f, 2].Style.Font.Bold = true;
                                }
                                CambioEstado = false;
                                Subestado = "";
                                CambioSubEstado = true;

                            } //  Final Testeo campo estado

                            //TesteoCampoSubestado
                            if (Subestado == "" || Subestado == rRS["subestado_global"].ToString())
                            {
                                Subestado = rRS["subestado_global"].ToString();
                                if (CambioSubEstado)
                                {
                                    workSheet.Cells[f, 3].Value = Subestado;
                                    workSheet.Cells[f, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                }
                                CambioSubEstado = false;
                            }
                            else
                            {
                                //Cambio Subestado Aumento fila
                                f = f + 1;
                                Subestado = rRS["subestado_global"].ToString();
                                CambioSubEstado = true;
                                if (CambioSubEstado)
                                {
                                    workSheet.Cells[f, 3].Value = Subestado;
                                    workSheet.Cells[f, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                }
                                CambioSubEstado = false;
                            } //  Final TesteoCampoSubestado

                            ////////Rellenamos Cuadro por columnas y vamos almacenando valores
                            DiaPintar = Convert.ToDateTime(rRS["fh_carga"]);
                            DiaAux = DiaPintar.ToString("dd/MM/yyyy");

                            for (int p = 4; p < 9; p++)
                            {     
                                if (DiaAux == workSheet.Cells[4, p].Text)
                                {
                                    // pinto conteo e importe por día
                                    if (rRS["conteo"].ToString()=="")
                                    {
                                        workSheet.Cells[f, p].Value =0;
                                    }
                                    else
                                    {
                                        workSheet.Cells[f, p].Value = Convert.ToInt32(rRS["conteo"].ToString());
                                    }     
                                    workSheet.Cells[f, p].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    if (rRS["importe"] is null || rRS["importe"].ToString() == "")
                                    {
                                        workSheet.Cells[f, p + 5].Value=0;  
                                    }
                                    else
                                    {
                                        workSheet.Cells[f, p + 5].Value = Convert.ToDouble(rRS["importe"]);
                                    }
                                    workSheet.Cells[f, p + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    
                                    //Voy acumulando los importes y los totales de los días
                                    if (DiaAux == Dia1)
                                    {
                                        TotalAgoraDia1 = TotalAgoraDia1 + Convert.ToInt32(rRS["conteo"]);

                                        if (rRS["importe"] is null || rRS["importe"].ToString()== "")
                                        {}
                                        else
                                        {
                                            TotalImporteDia1 = TotalImporteDia1 + Convert.ToDouble(rRS["importe"]);
                                            TotalImporteDia1Final = TotalImporteDia1Final + Convert.ToDouble(rRS["importe"]);
                                        }   
                                        TotalAgoraDia1Final= TotalAgoraDia1Final + Convert.ToInt32(rRS["conteo"]);   
                                        break;
                                    }
                                    if (DiaAux == Dia2)
                                    {
                                        TotalAgoraDia2 = TotalAgoraDia2 + Convert.ToInt32(rRS["conteo"]);

                                        if (rRS["importe"] is null || rRS["importe"].ToString() == "")
                                        { }
                                        else
                                        {
                                            TotalImporteDia2 = TotalImporteDia2 + Convert.ToDouble(rRS["importe"]);
                                            TotalImporteDia2Final = TotalImporteDia2Final + Convert.ToDouble(rRS["importe"]);
                                        }     
                                        TotalAgoraDia2Final = TotalAgoraDia2Final + Convert.ToInt32(rRS["conteo"]);                 
                                        break;
                                    }
                                    if (DiaAux == Dia3)
                                    {
                                        TotalAgoraDia3 = TotalAgoraDia3 + Convert.ToInt32(rRS["conteo"]);

                                        if (rRS["importe"] is null || rRS["importe"].ToString() == "")
                                        { }
                                        else
                                        {
                                            TotalImporteDia3 = TotalImporteDia3 + Convert.ToDouble(rRS["importe"]);
                                            TotalImporteDia3Final = TotalImporteDia3Final + Convert.ToDouble(rRS["importe"]);
                                        }         
                                        TotalAgoraDia3Final = TotalAgoraDia3Final + Convert.ToInt32(rRS["conteo"]);         
                                        break;
                                    }
                                    if (DiaAux == Dia4)
                                    {
                                        TotalAgoraDia4 = TotalAgoraDia4 + Convert.ToInt32(rRS["conteo"]);
                                        if (rRS["importe"] is null || rRS["importe"].ToString() == "")
                                        {                                        }
                                        else
                                        {
                                            TotalImporteDia4 = TotalImporteDia4 + Convert.ToDouble(rRS["importe"]);
                                            TotalImporteDia4Final = TotalImporteDia4Final + Convert.ToDouble(rRS["importe"]);
                                        }                               
                                        TotalAgoraDia4Final = TotalAgoraDia4Final + Convert.ToInt32(rRS["conteo"]);       
                                        break;
                                    }
                                    if (DiaAux == Dia5)
                                    {
                                        TotalAgoraDia5 = TotalAgoraDia5 + Convert.ToInt32(rRS["conteo"]);
                                        if (rRS["importe"] is null || rRS["importe"].ToString() == "")
                                        {                                        }
                                        else
                                        {
                                            TotalImporteDia5 = TotalImporteDia5 + Convert.ToDouble(rRS["importe"]);
                                            TotalImporteDia5Final = TotalImporteDia5Final + Convert.ToDouble(rRS["importe"]);
                                        }         
                                        TotalAgoraDia5Final = TotalAgoraDia5Final + Convert.ToInt32(rRS["conteo"]);     
                                        break;
                                    }
                                } // if (DiaAux == workSheet.Cells[4, p].Text)
                            } // Fin for (int p = 4; p < 9; p++)
                        }  // Final  while (rRS.Read())


                        //Pinto el total del subestado que no he pintado al terminar 
                        ///Ajustamos la caja con el texto
                        RangoInterno = "B" + InicioCaja.ToString() + ":B" + f.ToString();
                        workSheet.Cells[RangoInterno].Merge = true;
                        workSheet.Cells[RangoInterno].Style.WrapText = true;
                        workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        InicioCaja = f;

                        //Cambio Estado Aumento fila
                        f = f + 1;

                        //Pinto las lineas de totales en gris
                        workSheet.Cells[f, 3].Value = "Total " + Estado;
                        workSheet.Cells[f, 3].Style.Font.Bold = true;
                        //////RangoPintoGris = "B" + f.ToString() + ":C" + f.ToString();
                        //////workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        RangoPintoGris = "B" + f.ToString() + ":M" + f.ToString();
                        workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaGris);
                        //////workSheet.Cells[f, c].Value = TotalConteoResponsable;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, c].Style.Font.Bold = true;
                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        //Cambia el estado: Pego la suma de conteo y de importe de cada dia
                        for (b = 4; b < 9; b++)
                        {
                            switch (b)
                            {
                                case 4:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia1;
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia1;
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    TotalAgoraDia1 = 0;
                                    TotalImporteDia1 = 0;
                                    break;
                                case 5:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia2;
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia2;
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    TotalAgoraDia2 = 0;
                                    TotalImporteDia2 = 0;
                                    break;
                                case 6:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia3;
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia3;
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    TotalAgoraDia3 = 0;
                                    TotalImporteDia3 = 0;
                                    break;
                                case 7:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia4;
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia4;
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    TotalAgoraDia4 = 0;
                                    TotalImporteDia4 = 0;
                                    break;
                                case 8:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia5;
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia5;
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    TotalAgoraDia5 = 0;
                                    TotalImporteDia5 = 0;
                                    break;
                            }
                        }

                        //fin Pinto el total del subestado que no he pintado por el cambio de agora


                        ///Ajustamos la caja con el texto
                        if (FilaCambioAgora == 0)
                        {
                            RangoInterno = "A5:A" + f.ToString();
                        }
                        else {
                            RangoInterno = "A" + (FilaCambioAgora + 1).ToString() + ":A" + f.ToString();
                        }
                       
                        workSheet.Cells[RangoInterno].Merge = true;
                        workSheet.Cells[RangoInterno].Style.WrapText = true;
                        workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        //Ponemos a cero todas las cajas que esten sin rellenar
                        for (int p = 5; p < f+1; p++)
                        {
                            if (workSheet.Cells[p, 4].Text == "") {workSheet.Cells[p, 4].Value = 0; workSheet.Cells[p, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin); }
                            if (workSheet.Cells[p, 5].Text == "") { workSheet.Cells[p, 5].Value = 0; workSheet.Cells[p, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin); }
                            if (workSheet.Cells[p, 6].Text == "") { workSheet.Cells[p, 6].Value = 0; workSheet.Cells[p, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin); }
                            if (workSheet.Cells[p, 7].Text == "") { workSheet.Cells[p, 7].Value = 0; workSheet.Cells[p, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin); }
                            if (workSheet.Cells[p, 8].Text == "") { workSheet.Cells[p, 8].Value = 0; workSheet.Cells[p, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin); }
                            if (workSheet.Cells[p, 9].Text == "") { workSheet.Cells[p, 9].Value = 0; workSheet.Cells[p, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin); }
                            if (workSheet.Cells[p, 10].Text == "") { workSheet.Cells[p, 10].Value = 0; workSheet.Cells[p, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin); }
                            if (workSheet.Cells[p, 11].Text == "") { workSheet.Cells[p, 11].Value = 0; workSheet.Cells[p, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin); }
                            if (workSheet.Cells[p, 12].Text == "") { workSheet.Cells[p, 12].Value = 0; workSheet.Cells[p, 12].Style.Border.BorderAround(ExcelBorderStyle.Thin); }
                            if (workSheet.Cells[p, 13].Text == "") { workSheet.Cells[p, 13].Value = 0; workSheet.Cells[p, 13].Style.Border.BorderAround(ExcelBorderStyle.Thin); }
                        }

                        //////allCells = workSheet.Cells[1,1,f,13];
                        //////allCells.AutoFitColumns();
                        workSheet.Cells["A1:M" + f.ToString()].AutoFitColumns();

                        //Pintamos los TOTALES
                        //Pinto las lineas de totales AGORA 
                        f = f + 1;
                        if (Agora == "N")
                        {
                            workSheet.Cells[f, 3].Value = "Total No Ágora";
                            FilaCambioAgora = 1000; //Pongo 1000 para que no de error al restar el valor en el caso de que 
                            //sólo hay No Agora, filacambioagora=0 y da error Convert.ToInt32(workSheet.Cells[FilaCambioAgora, b].Value);
                        }
                        else
                        {
                            workSheet.Cells[f, 3].Value = "Total Si Ágora";
                        }
                        workSheet.Cells[f, 3].Style.Font.Bold = true;
                        //////RangoPintoGris = "A" + f.ToString() + ":C" + f.ToString();
                        //////workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        RangoPintoGris = "A" + f.ToString() + ":M" + f.ToString();
                        workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaTitulo);
                        workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells[RangoPintoGris].Style.Font.Bold = true;
                        //////workSheet.Cells[f, c].Value = TotalConteoResponsable;
                        //////workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        //////workSheet.Cells[f, c].Style.Font.Bold = true;
                        //////workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        for (b = 4; b < 9; b++)
                        {
                            switch (b)
                            {
                                case 4:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia1Final - Convert.ToInt32(workSheet.Cells[FilaCambioAgora, b].Value);
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia1Final - Convert.ToInt32(workSheet.Cells[FilaCambioAgora, b + 5].Value);
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b].Style.Font.Bold = true;          
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    break;
                                case 5:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia2Final - Convert.ToInt32(workSheet.Cells[FilaCambioAgora, b].Value); ;
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia2Final - Convert.ToInt32(workSheet.Cells[FilaCambioAgora, b + 5].Value);
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    break;
                                case 6:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia3Final - Convert.ToInt32(workSheet.Cells[FilaCambioAgora, b].Value); ;
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia3Final - Convert.ToInt32(workSheet.Cells[FilaCambioAgora, b + 5].Value);
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    break;
                                case 7:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia4Final - Convert.ToInt32(workSheet.Cells[FilaCambioAgora, b].Value); ;
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia4Final - Convert.ToInt32(workSheet.Cells[FilaCambioAgora, b + 5].Value);
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    break;
                                case 8:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia5Final - Convert.ToInt32(workSheet.Cells[FilaCambioAgora, b].Value);
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia5Final - Convert.ToInt32(workSheet.Cells[FilaCambioAgora, b + 5].Value);
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    break;
                            }
                        }

                        //Pinto las lineas de total General 
                        f = f + 1;
                        workSheet.Cells[f, 3].Value = "Total General";
                        workSheet.Cells[f, 3].Style.Font.Bold = true;
                        //////RangoPintoGris = "A" + f.ToString() + ":C" + f.ToString();
                        //////workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        RangoPintoGris = "A" + f.ToString() + ":M" + f.ToString();
                        workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaCabecera);
                        workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells[RangoPintoGris].Style.Font.Bold = true;
                        //////workSheet.Cells[f, c].Value = TotalConteoResponsable;
                        //////workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                        //////workSheet.Cells[f, c].Style.Font.Bold = true;
                        //////workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        for (b = 4; b < 9; b++)
                        {
                            switch (b)
                            {
                                case 4:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia1Final;
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b+5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia1Final;
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    break;
                                case 5:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia2Final;
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia2Final;
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    break;
                                case 6:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia3Final;
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia3Final;
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    break;
                                case 7:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia4Final;
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia4Final;
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    break;
                                case 8:
                                    workSheet.Cells[f, b].Value = TotalAgoraDia5Final;
                                    workSheet.Cells[f, b + 5].Value = TotalImporteDia5Final;
                                    workSheet.Cells[f, b].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[f, b].Style.Font.Bold = true;
                                    workSheet.Cells[f, b + 5].Style.Font.Bold = true;
                                    break;
                            }
                        }


                        workSheet.Column(1).Width = 14;
                        workSheet.Column(2).Width = 28;
                        //////allCells = workSheet.Cells[1,1,f,13];
                        //////allCells.AutoFitColumns();
                        ////workSheet.Cells["A1:M" + f.ToString()].AutoFitColumns();

                        dbRS.CloseConnection();
                    }

                    //Muevo hoja cuatro después de la hoja uno
                    excelPackage.Workbook.Worksheets.MoveAfter(excelPackage.Workbook.Worksheets[3].Name, excelPackage.Workbook.Worksheets[0].Name);
                    //Muevo hoja cinco después de la hoja tres
                    excelPackage.Workbook.Worksheets.MoveAfter(excelPackage.Workbook.Worksheets[4].Name, excelPackage.Workbook.Worksheets[2].Name);


                    // FIN 28/04/2025 CREAMOS HOJAS/CUADROS RESUMEN a PARTIR DE TABLA  facturacionb2b_owner.sap_kee_resumen_es_por

                    //creamos las hojas de historico (5 últimos días) a partir de facturacionb2b_owner.sap_kee_resumen_es_por
                    for (int p = 1; p < 4; p++)
                    {
                        if (p == 1)
                        {
                            workSheet = excelPackage.Workbook.Worksheets.Add("HISTORICO ES");
                        }
                        if (p == 2)
                        {
                            workSheet = excelPackage.Workbook.Worksheets.Add("HISTORICO POR MT-BTE");
                        }
                        if (p == 3)
                        {
                            workSheet = excelPackage.Workbook.Worksheets.Add("HISTORICO POR BTN");
                        }

                        headerCells = workSheet.Cells[1, 1, 1, 32];
                        headerFont = headerCells.Style.Font;
                        //Ponemos columnas directamente
                        f = 1;
                        c = 1;

                        if (p == 1)
                        {
                            workSheet.Cells[f, c].Value = "CUPS20;PERIODO; MES;FH_DESDE;FH_HASTA;AREA_ESTADO;Subestado global;ESTADO GLOBAL;ESTADO GLOBAL A REPORTAR;IMPORTE PDTE FACTURAR;MESES PDTES FACTURAR;ÁGORA;CLIENTE;DIAS_ESTADO_SAP;DIAS_ESTADO_KEE;DIAS_ESTADO_GLOBAL;Incidencia_Facturacion;Estado_FAC_SE;Titulo_FAC;Incidencia_Medida;Reincidente;Estado incidencia;Fecha_Apertura;Prioridad;Titulo;E_S_Estado;FH_ALTA_SALESFORCE;FH_ALTA_KEE;FH_ALTA_SAP;FH_BAJA_SALESFORCE;Fecha BAJA KEE;FH_BAJA_SAP;EMPRESA;TIPO CLIENTE;NIF;FPSERCON;Nº INSTALACIÓN;TARIFA;CONTRATO;DISTRIBUIDORA;ESTADO;SUBESTADO;TAM;ULT FH DESDE FACTURADA;ULT FH HASTA FACTURADA;Estado periodo KEE;Área responsable KEE;Subestado KEE;Estado_KEE;FH_DESDE_KEE;FH_HASTA_KEE;Multipunto;Discrepancias;FH_CARGA";
                        }
                        if (p == 2)
                        {
                            workSheet.Cells[f, c].Value = "CUPS20;PERIODO; MES;FH_DESDE;FH_HASTA;AREA_ESTADO;Subestado global;ESTADO GLOBAL;ESTADO GLOBAL A REPORTAR;IMPORTE PDTE FACTURAR;MESES PDTES FACTURAR;ÁGORA;CLIENTE;DIAS_ESTADO_SAP;DIAS_ESTADO_KEE;DIAS_ESTADO_GLOBAL;Incidencia_Facturacion;Estado_FAC_SE;Titulo_FAC;Incidencia_Medida;Reincidente;Estado incidencia;Fecha_Apertura;Prioridad;Titulo;E_S_Estado;FH_ALTA_SALESFORCE;FH_ALTA_KEE;FH_ALTA_SAP;FH_BAJA_SALESFORCE;Fecha BAJA KEE;FH_BAJA_SAP;EMPRESA;SEGMENTO;TIPO CLIENTE;NIF;FPSERCON;Nº INSTALACIÓN;TARIFA;CONTRATO;DISTRIBUIDORA;ESTADO;SUBESTADO;TAM;ULT FH DESDE FACTURADA;ULT FH HASTA FACTURADA;Estado periodo KEE;Área responsable KEE;Subestado KEE;Estado_KEE;FH_DESDE_KEE;FH_HASTA_KEE;Multipunto;Discrepancias;FH_CARGA";
                        }
                        if (p == 3)
                        {
                            workSheet.Cells[f, c].Value = "CUPS20;PERIODO; MES;FH_DESDE;FH_HASTA;AREA_ESTADO;Subestado global;ESTADO GLOBAL;ESTADO GLOBAL A REPORTAR;IMPORTE PDTE FACTURAR;MESES PDTES FACTURAR;ÁGORA;CLIENTE;DIAS_ESTADO_SAP;DIAS_ESTADO_KEE;DIAS_ESTADO_GLOBAL;Incidencia_Facturacion;Estado_FAC_SE;Titulo_FAC;Incidencia_Medida;Reincidente;Estado incidencia;Fecha_Apertura;Prioridad;Titulo;E_S_Estado;FH_ALTA_SALESFORCE;FH_ALTA_KEE;FH_ALTA_SAP;FH_BAJA_SALESFORCE;Fecha BAJA KEE;FH_BAJA_SAP;EMPRESA;SEGMENTO;TIPO CLIENTE;NIF;FPSERCON;Nº INSTALACIÓN;TARIFA;CONTRATO;DISTRIBUIDORA;ESTADO;SUBESTADO;TAM;ULT FH DESDE FACTURADA;ULT FH HASTA FACTURADA;Estado periodo KEE;Área responsable KEE;Subestado KEE;Estado_KEE;FH_DESDE_KEE;FH_HASTA_KEE;Discrepancias;Nº Días PTE KEE;FH_CARGA";
                        }

                        headerCells = workSheet.Cells[1, 1, 1, c];
                        headerFont = headerCells.Style.Font;
                        headerFont.Bold = true;

                        if (p == 1)
                        {

                              //////strSql = "SELECT concat(CUPS20 ,';',PERIODO ,';',MES,';',if (FH_DESDE = '0000-00-00', null, FH_DESDE),';',if (FH_HASTA = '0000-00-00', null, FH_HASTA), ';' ,AREA_ESTADO, ';' ,Subestado_global , ';' , ESTADO_GLOBAL , ';' , ESTADO_GLOBAL_A_REPORTAR, ';' ,  IFNULL(replace(IMPORTE_PDTE_FACTURAR,'.',','), '')  , ';' , MESES_PDTES_FACTURAR, ';' , Agora , ';' , Cliente , ';' , DIAS_ESTADO_SAP, ';' , DIAS_ESTADO_GLOBAL , ';',Incidencia_Facturacion , ';', Estado_FAC_SE , ';', ';' , Titulo_FAC , ';' , Incidencia_Medida , ';', Reincidente , ';' , Estado_incidencia , ';', IFNULL(Fecha_Apertura, '') , ';' , Prioridad , ';' , Titulo , ';' , E_S_Estado , ';' , IFNULL(FH_ALTA_SALESFORCE, ''), ';', IFNULL(Fecha_ALTA_KEE, '') , ';' , ifnull(FH_ALTA_SAP, '') , ';' , IFNULL(FH_BAJA_SALESFORCE, '') , ';', IFNULL(Fecha_BAJA_KEE, '') , ';' , IFNULL(FH_BAJA_SAP, '') , ';' ,Empresa , ';' ,  tipo_cliente , ';' , nif , ';' , FPSERCON , ';' , N_INSTALACION , ';' , tarifa , ';' , contrato , ';' , Distribuidora , ';' ,   Estado , ';' , Subestado , ';' , IFNULL(replace(TAM,'.',','), '')  , ';' , IFNULL(ULT_FH_DESDE_FACTURADA,'') , ';', IFNULL(ULT_FH_HASTA_FACTURADA,'') , ';' , Estado_periodo_KEE , ';' , Area_responsable_KEE , ';' , Subestado_KEE , ';' , replace(Estado_KEE,';','-') , ';', ifnull(if (FH_DESDE_KEE = '0000-00-00', null, FH_DESDE_KEE),''), ';' , ifnull(if (FH_HASTA_KEE = '0000-00-00', null, FH_HASTA_KEE),''), ';' , Multipunto , ';' , replace(Discrepancia,';','-') , ';' ,  FH_CARGA  ) as Cadena "
                              //////+ " FROM sap_KEE_Resumen_ES_POR_MT_BTE_New "
                              //////+ " Where (segmento='' or segmento IS NULL)"
                              //////+ " AND FH_CARGA >= ("
                              //////+ "   SELECT MIN(A.FH_CARGA) FROM"
                              //////+ "   ("
                              //////+ "       SELECT distinct FH_CARGA"
                              //////+ "       FROM sap_KEE_Resumen_ES_POR_MT_BTE_New"
                              //////+ "       Where (segmento = '' or segmento is null)"
                              //////+ "      ORDER BY FH_CARGA DESC"
                              //////+ "       LIMIT 5"
                              //////+ "   ) AS A"
                              //////+ ")"
                              //////+ " order by FH_CARGA desc";

                            strSql = "SELECT nvl(CUPS20, '') || ';' || nvl(PERIODO, '') || ';' || nvl(MES, '') || ';' || nvl(to_char(FH_DESDE, 'yyyy-mm-dd'), '') " 
                                + " || ';'|| nvl(to_char(FH_HASTA, 'yyyy-mm-dd'), '') || ';' || nvl(AREA_ESTADO, '') || ';' || nvl(Subestado_global, '') || ';' || nvl(ESTADO_GLOBAL, '') || ';' || nvl(ESTADO_GLOBAL_A_REPORTAR, '') "
                                + " || ';' || nvl(IMPORTE_PDTE_FACTURAR, 0) || ';' || nvl(MESES_PDTES_FACTURAR, 0) || ';' || nvl(Agora, '') || ';' || replace(nvl(Cliente, ''),';',' ') "
                                + " || ';' || nvl(DIAS_ESTADO_SAP, '') || ';' || nvl(DIAS_ESTADO_KEE, '') || ';' || nvl(DIAS_ESTADO_GLOBAL, '') || ';' || nvl(Incidencia_Facturacion, '') "
                                + " || ';' || nvl(Estado_FAC_SE, '') || ';' || nvl(Titulo_FAC, '') || ';' || nvl(Incidencia_Medida, '') || ';' || nvl(Reincidente, '') "
                                + " || ';' || nvl(Estado_incidencia, '') || ';' || nvl(Fecha_Apertura, '') || ';' || nvl(Prioridad, '') || ';' || nvl(Titulo, '') "
                                + " || ';' || nvl(E_S_Estado, '') || ';' || nvl(to_char(FH_ALTA_SALESFORCE, 'yyyy-mm-dd'), '') "
                                + " || ';' || nvl(to_char(Fecha_ALTA_KEE, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(FH_ALTA_SAP, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(FH_BAJA_SALESFORCE, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(Fecha_BAJA_KEE, 'yyyy-mm-dd'), '') "
                                + " || ';' || nvl(to_char(FH_BAJA_SAP, 'yyyy-mm-dd'), '') || ';' || nvl(Empresa, '') "
                                + " || ';' || nvl(tipo_cliente, '') || ';' || nvl(nif, '') || ';' || nvl(to_char(FPSERCON, 'yyyy-mm-dd'), '') "
                                + " || ';' || nvl(N_INSTALACION, '') || ';' || nvl(tarifa, '') || ';' || nvl(contrato, '') || ';' || nvl(Distribuidora, '') || ';' || nvl(Estado, '') || ';' || nvl(Subestado, '') "
                                + " || ';' || nvl(TAM, 0) || ';' || nvl(to_char(ULT_FH_DESDE_FACTURADA, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(ULT_FH_HASTA_FACTURADA, 'yyyy-mm-dd'), '') "
                                + " || ';' || nvl(Estado_periodo_KEE, '') || ';' || nvl(Area_responsable_KEE, '') || ';' || nvl(Subestado_KEE, '') || ';' || replace(Estado_KEE, ';', '-') "
                                + " || ';' || nvl(to_char(FH_DESDE_KEE, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(FH_HASTA_KEE, 'yyyy-mm-dd'), '') || ';' || nvl(Multipunto, '') || ';' || replace(nvl(Discrepancia, ''), ';', '-') || ';' || FH_CARGA "
                                + " as Cadena"
                                + " FROM facturacionb2b_owner.sap_kee_resumen_es_por"
                                + " Where(segmento = '' or segmento IS NULL)"
                                + " AND FH_CARGA >= ("
                                + "  SELECT MIN(A.FH_CARGA) FROM"
                                + "  ("
                                + "    SELECT distinct FH_CARGA"
                                + "    FROM facturacionb2b_owner.sap_kee_resumen_es_por"
                                + "    Where(segmento = '' or segmento is null)"
                                + "    ORDER BY FH_CARGA DESC"
                                + "    LIMIT 5"
                                + "  ) AS A"
                                + " )"
                                + " order by FH_CARGA desc";
                        }

                        if (p == 2)
                        {

                            //////strSql = "SELECT concat(CUPS20 ,';',PERIODO ,';',MES,';',if (FH_DESDE = '0000-00-00', null, FH_DESDE),';',if (FH_HASTA = '0000-00-00', null, FH_HASTA), ';' ,AREA_ESTADO, ';' ,Subestado_global , ';' , ESTADO_GLOBAL , ';' , ESTADO_GLOBAL_A_REPORTAR, ';' , IFNULL(replace(IMPORTE_PDTE_FACTURAR,'.',','), '')  , ';' , MESES_PDTES_FACTURAR, ';' , Agora , ';' , Cliente , ';' , DIAS_ESTADO_SAP, ';' , DIAS_ESTADO_KEE, ';', DIAS_ESTADO_GLOBAL , ';',Incidencia_Facturacion , ';', Estado_FAC_SE , ';',  Titulo_FAC , ';' , Incidencia_Medida , ';', Reincidente , ';' , Estado_incidencia , ';', IFNULL(Fecha_Apertura, '') , ';' , Prioridad , ';' , Titulo , ';' , E_S_Estado , ';' , IFNULL(FH_ALTA_SALESFORCE, ''), ';', IFNULL(Fecha_ALTA_KEE, '') , ';' , ifnull(FH_ALTA_SAP, '') , ';' , IFNULL(FH_BAJA_SALESFORCE, '') , ';', IFNULL(Fecha_BAJA_KEE, '') , ';' , IFNULL(FH_BAJA_SAP, '') , ';' ,Empresa , ';' , SEGMENTO, ';' ,  tipo_cliente , ';' , nif , ';' , FPSERCON , ';' , N_INSTALACION , ';' , tarifa , ';' , contrato , ';' , Distribuidora , ';' ,   Estado , ';' , Subestado , ';' , IFNULL(replace(TAM,'.',','), '')  , ';' , IFNULL(ULT_FH_DESDE_FACTURADA,''), ';',IFNULL(ULT_FH_HASTA_FACTURADA,'') , ';' , Estado_periodo_KEE , ';' , Area_responsable_KEE , ';' , Subestado_KEE , ';' , replace(Estado_KEE,';','-') , ';', ifnull(if (FH_DESDE_KEE = '0000-00-00', null, FH_DESDE_KEE),''), ';' , ifnull(if (FH_HASTA_KEE = '0000-00-00', null, FH_HASTA_KEE),''), ';' , Multipunto , ';' , replace(Discrepancia,';','-') , ';' ,  FH_CARGA  ) as Cadena "
                            //////+ " FROM sap_KEE_Resumen_ES_POR_MT_BTE_New "
                            //////+ " Where segmento in ('MT', 'BTE') "
                            //////+ " AND FH_CARGA >= ("
                            //////+ "   SELECT MIN(A.FH_CARGA) FROM"
                            //////+ "   ("
                            //////+ "       SELECT distinct FH_CARGA"
                            //////+ "       FROM sap_KEE_Resumen_ES_POR_MT_BTE_New"
                            //////+ "       Where segmento in ('MT', 'BTE')"
                            //////+ "      ORDER BY FH_CARGA DESC"
                            //////+ "       LIMIT 5"
                            //////+ "   ) AS A"
                            //////+ ")"
                            //////+ " order by FH_CARGA desc";
                            ///
                            /// 
                            ////"CUPS20;PERIODO; MES;FH_DESDE;FH_HASTA;AREA_ESTADO;Subestado global;ESTADO GLOBAL;ESTADO GLOBAL A REPORTAR;IMPORTE PDTE FACTURAR;MESES PDTES FACTURAR;ÁGORA;CLIENTE;
                            ///DIAS_ESTADO_SAP;DIAS_ESTADO_KEE;DIAS_ESTADO_GLOBAL;Incidencia_Facturacion;Estado_FAC_SE;Titulo_FAC;Incidencia_Medida;Reincidente;Estado incidencia;Fecha_Apertura;
                            ///Prioridad;Titulo;E_S_Estado;FH_ALTA_SALESFORCE;FH_ALTA_KEE;FH_ALTA_SAP;FH_BAJA_SALESFORCE;Fecha BAJA KEE;FH_BAJA_SAP;EMPRESA;SEGMENTO;TIPO CLIENTE;NIF;FPSERCON;Nº INSTALACIÓN;TARIFA;CONTRATO;DISTRIBUIDORA;ESTADO;SUBESTADO;TAM;ULT FH DESDE FACTURADA;ULT FH HASTA FACTURADA;Estado periodo KEE;Área responsable KEE;Subestado KEE;Estado_KEE;FH_DESDE_KEE;FH_HASTA_KEE;Multipunto;Discrepancias;FH_CARGA";

                            strSql = "SELECT nvl(CUPS20, '') || ';' || nvl(PERIODO, '') || ';' || nvl(MES, '') || ';' || nvl(to_char(FH_DESDE, 'yyyy-mm-dd'), '') "
                                + " || ';'|| nvl(to_char(FH_HASTA, 'yyyy-mm-dd'), '') || ';' || nvl(AREA_ESTADO, '') || ';' || nvl(Subestado_global, '') || ';' || nvl(ESTADO_GLOBAL, '') || ';' || nvl(ESTADO_GLOBAL_A_REPORTAR, '') "
                                + " || ';' || nvl(IMPORTE_PDTE_FACTURAR, 0) || ';' || nvl(MESES_PDTES_FACTURAR, 0) || ';' || nvl(Agora, '') || ';' || replace(nvl(Cliente, ''),';',' ') "
                                + " || ';' || nvl(DIAS_ESTADO_SAP, '') || ';' || nvl(DIAS_ESTADO_KEE, '') || ';' ||nvl(DIAS_ESTADO_GLOBAL, '') || ';' || nvl(Incidencia_Facturacion, '') "
                                + " || ';' || nvl(Estado_FAC_SE, '') || ';' || nvl(Titulo_FAC, '') || ';' || nvl(Incidencia_Medida, '') || ';' || nvl(Reincidente, '') "
                                + " || ';' || nvl(Estado_incidencia, '') || ';' || nvl(Fecha_Apertura, '') || ';' || nvl(Prioridad, '') || ';' || nvl(Titulo, '') "
                                + " || ';' || nvl(E_S_Estado, '') || ';' || nvl(to_char(FH_ALTA_SALESFORCE, 'yyyy-mm-dd'), '') "
                                + " || ';' || nvl(to_char(Fecha_ALTA_KEE, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(FH_ALTA_SAP, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(FH_BAJA_SALESFORCE, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(Fecha_BAJA_KEE, 'yyyy-mm-dd'), '') "
                                + " || ';' || nvl(to_char(FH_BAJA_SAP, 'yyyy-mm-dd'), '') || ';' || nvl(Empresa, '') || ';' || nvl(segmento, '')"
                                + " || ';' || nvl(tipo_cliente, '') || ';' || nvl(nif, '') || ';' || nvl(to_char(FPSERCON, 'yyyy-mm-dd'), '') "
                                + " || ';' || nvl(N_INSTALACION, '') || ';' || nvl(tarifa, '') || ';' || nvl(contrato, '') || ';' || nvl(Distribuidora, '') || ';' || nvl(Estado, '') || ';' || nvl(Subestado, '') "
                                + " || ';' || nvl(TAM, 0) || ';' || nvl(to_char(ULT_FH_DESDE_FACTURADA, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(ULT_FH_HASTA_FACTURADA, 'yyyy-mm-dd'), '') "
                                + " || ';' || nvl(Estado_periodo_KEE, '') || ';' || nvl(Area_responsable_KEE, '') || ';' || nvl(Subestado_KEE, '') || ';' || replace(Estado_KEE, ';', '-') "
                                + " || ';' || nvl(to_char(FH_DESDE_KEE, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(FH_HASTA_KEE, 'yyyy-mm-dd'), '') || ';' || nvl(Multipunto, '') || ';' || replace(nvl(Discrepancia, ''), ';', '-') || ';' || FH_CARGA "
                                + " as Cadena"
                                + " FROM facturacionb2b_owner.sap_kee_resumen_es_por"
                                + " Where segmento in ('MT', 'BTE','AT','MAT') "
                                + " AND FH_CARGA >= ("
                                + "  SELECT MIN(A.FH_CARGA) FROM"
                                + "  ("
                                + "    SELECT distinct FH_CARGA"
                                + "    FROM facturacionb2b_owner.sap_kee_resumen_es_por"
                                + "    Where segmento in ('MT', 'BTE','AT','MAT') "
                                + "    ORDER BY FH_CARGA DESC"
                                + "    LIMIT 5"
                                + "  ) AS A"
                                + " )"
                                + " order by FH_CARGA desc";
                        }
                   
                        if (p == 3)
                        {

                            //////strSql = "SELECT concat(CUPS20 ,';',PERIODO ,';',MES,';',if (FH_DESDE = '0000-00-00', null, FH_DESDE),';',if (FH_HASTA = '0000-00-00', null, FH_HASTA), ';' ,AREA_ESTADO, ';' ,Subestado_global , ';' , ESTADO_GLOBAL , ';' , ESTADO_GLOBAL_A_REPORTAR, ';' , IFNULL(replace(IMPORTE_PDTE_FACTURAR,'.',','), '')  , ';' , MESES_PDTES_FACTURAR, ';' , Agora , ';' , Cliente , ';' , DIAS_ESTADO_SAP, ';' , DIAS_ESTADO_KEE, ';', DIAS_ESTADO_GLOBAL , ';',Incidencia_Facturacion , ';', Estado_FAC_SE , ';' , Titulo_FAC , ';' , Incidencia_Medida , ';', Reincidente , ';' , Estado_incidencia , ';', IFNULL(Fecha_Apertura, '') , ';' , Prioridad , ';' , Titulo , ';' , E_S_Estado , ';' , IFNULL(FH_ALTA_SALESFORCE, ''), ';', IFNULL(Fecha_ALTA_KEE, '') , ';' , ifnull(FH_ALTA_SAP, '') , ';' , IFNULL(FH_BAJA_SALESFORCE, '') , ';', IFNULL(Fecha_BAJA_KEE, '') , ';' , IFNULL(FH_BAJA_SAP, '') , ';' ,Empresa , ';' , SEGMENTO, ';' ,  tipo_cliente , ';' , nif , ';' , FPSERCON , ';' , N_INSTALACION , ';' , tarifa , ';' , contrato , ';' , Distribuidora , ';' ,   Estado , ';' , Subestado , ';' , IFNULL(replace(TAM,'.',','), '') , ';' , ifnull(ULT_FH_DESDE_FACTURADA,'') , ';',ifnull(ULT_FH_HASTA_FACTURADA,'') , ';' , Estado_periodo_KEE , ';' , Area_responsable_KEE , ';' , Subestado_KEE , ';' , replace(Estado_KEE,';','-') , ';', ifnull(if (FH_DESDE_KEE = '0000-00-00', null, FH_DESDE_KEE),''), ';' , ifnull(if (FH_HASTA_KEE = '0000-00-00', null, FH_HASTA_KEE),''), ';'  , replace(Discrepancia,';','-') , ';' , ifnull(NDIAS_PTE_KEE,''), ';', FH_CARGA  ) as Cadena "
                            //////  + " FROM sap_KEE_Resumen_ES_POR_MT_BTE_New "
                            //////  + " Where segmento ='BTN' "
                            //////  + " AND FH_CARGA >= ("
                            //////  + "   SELECT MIN(A.FH_CARGA) FROM"
                            //////  + "   ("
                            //////  + "       SELECT distinct FH_CARGA"
                            //////  + "       FROM sap_KEE_Resumen_ES_POR_MT_BTE_New"
                            //////  + "       Where segmento in ('BTN')"
                            //////  + "      ORDER BY FH_CARGA DESC"
                            //////  + "       LIMIT 5"
                            //////  + "   ) AS A"
                            //////  + ")"
                            //////  + " order by FH_CARGA desc";
                            ///
                            strSql = "SELECT nvl(CUPS20, '') || ';' || nvl(PERIODO, '') || ';' || nvl(MES, '') || ';' || nvl(to_char(FH_DESDE, 'yyyy-mm-dd'), '') "
                               + " || ';'|| nvl(to_char(FH_HASTA, 'yyyy-mm-dd'), '') || ';' || nvl(AREA_ESTADO, '') || ';' || nvl(Subestado_global, '') || ';' || nvl(ESTADO_GLOBAL, '') || ';' || nvl(ESTADO_GLOBAL_A_REPORTAR, '') "
                               + " || ';' || nvl(IMPORTE_PDTE_FACTURAR, 0) || ';' || nvl(MESES_PDTES_FACTURAR, 0) || ';' || nvl(Agora, '') || ';' || replace(nvl(Cliente, ''),';',' ') "
                               + " || ';' || nvl(DIAS_ESTADO_SAP, '') || ';' || nvl(DIAS_ESTADO_KEE, '') || ';' || nvl(DIAS_ESTADO_GLOBAL, '') || ';' || nvl(Incidencia_Facturacion, '') "
                               + " || ';' || nvl(Estado_FAC_SE, '') || ';' || nvl(Titulo_FAC, '') || ';' || nvl(Incidencia_Medida, '') || ';' || nvl(Reincidente, '') "
                               + " || ';' || nvl(Estado_incidencia, '') || ';' || nvl(Fecha_Apertura, '') || ';' || nvl(Prioridad, '') || ';' || nvl(Titulo, '') "
                               + " || ';' || nvl(E_S_Estado, '') || ';' || nvl(to_char(FH_ALTA_SALESFORCE, 'yyyy-mm-dd'), '') "
                               + " || ';' || nvl(to_char(Fecha_ALTA_KEE, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(FH_ALTA_SAP, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(FH_BAJA_SALESFORCE, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(Fecha_BAJA_KEE, 'yyyy-mm-dd'), '') "
                               + " || ';' || nvl(to_char(FH_BAJA_SAP, 'yyyy-mm-dd'), '') || ';' || nvl(Empresa, '') || ';' || nvl(segmento, '')"
                               + " || ';' || nvl(tipo_cliente, '') || ';' || nvl(nif, '') || ';' || nvl(to_char(FPSERCON, 'yyyy-mm-dd'), '') "
                               + " || ';' || nvl(N_INSTALACION, '') || ';' || nvl(tarifa, '') || ';' || nvl(contrato, '') || ';' || nvl(Distribuidora, '') || ';' || nvl(Estado, '') || ';' || nvl(Subestado, '') "
                               + " || ';' || nvl(TAM, 0) || ';' || nvl(to_char(ULT_FH_DESDE_FACTURADA, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(ULT_FH_HASTA_FACTURADA, 'yyyy-mm-dd'), '') "
                               + " || ';' || nvl(Estado_periodo_KEE, '') || ';' || nvl(Area_responsable_KEE, '') || ';' || nvl(Subestado_KEE, '') || ';' || replace(Estado_KEE, ';', '-') "
                               + " || ';' || nvl(to_char(FH_DESDE_KEE, 'yyyy-mm-dd'), '') || ';' || nvl(to_char(FH_HASTA_KEE, 'yyyy-mm-dd'), '')  || ';' || replace(nvl(Discrepancia, ''), ';', '-') || ';' || nvl(ndias_pte_kee, 0) || ';' || FH_CARGA "
                               + " as Cadena"
                               + " FROM facturacionb2b_owner.sap_kee_resumen_es_por"
                                + " Where segmento ='BTN' "
                                + " AND FH_CARGA >= ("
                                + "  SELECT MIN(A.FH_CARGA) FROM"
                                + "  ("
                                + "    SELECT distinct FH_CARGA"
                                + "    FROM facturacionb2b_owner.sap_kee_resumen_es_por"
                                + "    Where segmento ='BTN' "
                                + "    ORDER BY FH_CARGA DESC"
                                + "    LIMIT 5"
                                + "  ) AS A"
                                + " )"
                                + " order by FH_CARGA desc";
                        }

                        //////db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        //////command = new MySqlCommand(strSql, db.con);
                        //////r = command.ExecuteReader();
                        ///
                        dbRS = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD_FACTURACION);
                        commandRS = new OdbcCommand(strSql, dbRS.con);
                        rRS = commandRS.ExecuteReader();
                       
                        f = 1;

                        List<string> lista_Prueba = new List<string>();

                        while (rRS.Read())
                        {
                            f++;
                            lista_Prueba.Add(ArreglaCadena(rRS["Cadena"].ToString()));
    
                        }

                        if (p == 1)
                        {
                            TotalPruebaES = f;
                        }
                        if (p == 2)
                        {
                            TotalPruebaMTBTE = f;
                        }
                        if (p == 3)
                        {
                            TotalPruebaBTN = f;
                        }

                        workSheet.Cells["A2"].LoadFromCollection(lista_Prueba);

                        //////db.CloseConnection();
                        dbRS.CloseConnection();
                    }
                    /// Fin Prueba columnas2

                    //////CREAMOS LAS  HOJAS RESUMEN EN TABLAS DINAMICAS *********************************************************************************
                    for (int k = 1; k < 4; k++)
                    {

                        TotalParcial = 0;
                        TotalCuadro = 0;
                        TotalGeneral = 0;
                        CuadroInicio = 0;
                        CuadroFin = 0;
                        ExistenDatosCuadro = false;

                        if (k == 1)
                        {
                            workSheet = excelPackage.Workbook.Worksheets.Add("TD_ESP_MT");
                        }
                        if (k == 2)
                        {
                            workSheet = excelPackage.Workbook.Worksheets.Add("TD_PTG_MT+BTE");
                        }
                        if (k == 3)
                        {
                            workSheet = excelPackage.Workbook.Worksheets.Add("TD_PTG_BTN");
                        }

                        //////headerCells = workSheet.Cells[1, 1, 1, 32];
                        //////headerFont = headerCells.Style.Font;

                        //////for (int t = 1; t < 3; t++)
                        //////{

                        PintarEstado = "";
                        PintarSubestado = "";
                        TotalParcial = 0;
                        TotalCuadro = 0;
                        TotalGeneral = 0;
                        CuadroFin = 0;

                        headerCells = workSheet.Cells[1, 1, 1, 1];
                        headerFont = headerCells.Style.Font;
                        headerFont.Bold = true;

                        f = 2;
                        c = 2;

                        workSheet.Cells[f, c].Value = "ESTADO_GLOBAL"; c++;
                        workSheet.Cells[f, c].Value = "SUBESTADO"; c++;
                        workSheet.Cells[f, c].Value = "INCIDENCIA_FACTURACION"; c++;
                        workSheet.Cells[f, c].Value = "INCIDENCIA_MEDIDA"; c++;
                        workSheet.Cells[f, c].Value = "TOTAL"; c++;

                        headerCells = workSheet.Cells[f, 2, f, c];
                        headerFont = headerCells.Style.Font;
                        headerFont.Bold = true;
                        BordeCuadro = "B2:F2";
                        workSheet.Cells[BordeCuadro].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        CuadroInicio = 3;                        

                        if (k == 1) 
                            {
                                    //////strSql = "SELECT estado_global,  subestado, Incidencia_Facturacion, Incidencia_Medida, sum(CASE WHEN periodo = '1P' then 1 ELSE '' END) AS 'P1', sum(CASE WHEN periodo = '>1P' then 1 ELSE '' END) AS '>1P', COUNT(subestado) AS Total"
                                    //////    + " from sap_KEE_Resumen_ES_POR_MT_BTE_New "
                                    //////    //////+ " INNER JOIN SAP_PILOTO "
                                    //////    //////+ " ON sap_KEE_Resumen_ES_POR_MT_BTE_New.CUPS20 = SAP_PILOTO.CD_CUPS "
                                    //////    + " WHERE fh_carga = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
                                    //////    + " AND segmento IS NULL"
                                    //////    + " GROUP BY   estado_global,   subestado, Incidencia_Facturacion, Incidencia_Medida"
                                    //////    + " ORDER BY ESTADO_GLOBAL, subestado, Incidencia_Facturacion, Incidencia_Medida, PERIODO";

                                strSql = "select A.estado_global,  A.subestado, A.Incidencia_Facturacion, A.Incidencia_Medida,"
                                    + " sum(A.Puno) as P1, sum(A.MayorPuno) as P2, sum(A.Tota) as Total"
                                    + " from"
                                    + " ("
                                    + " SELECT estado_global, subestado, Incidencia_Facturacion, Incidencia_Medida,"
                                    + " sum(case when periodo = '1P' then 1 else 0 end) AS Puno,"
                                    + " sum(case when periodo = '>1P' then 1 else 0 end) AS MayorPuno,"
                                    + " count(subestado) AS Tota"
                                    + " from facturacionb2b_owner.sap_kee_resumen_es_por"
                                    + " WHERE fh_carga = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
                                    + " AND segmento IS NULL"
                                    + " GROUP BY   estado_global, subestado, Incidencia_Facturacion, Incidencia_Medida, periodo"
                                    + " ) as A"
                                    + " group by A.estado_global,  A.subestado, A.Incidencia_Facturacion, A.Incidencia_Medida"
                                    + " ORDER BY A.estado_global, A.subestado, A.Incidencia_Facturacion, A.Incidencia_Medida ";

                            }
                        if (k == 2)
                            {
                            //////strSql = "SELECT estado_global,  subestado,Incidencia_Facturacion,Incidencia_Medida,  sum(CASE WHEN periodo = '1P' then 1 ELSE '' END) AS 'P1', sum(CASE WHEN periodo = '>1P' then 1 ELSE '' END) AS '>1P', COUNT(subestado) AS Total"
                            //////    + " from sap_KEE_Resumen_ES_POR_MT_BTE_New "
                            //////    //////+ " INNER JOIN SAP_PILOTO "
                            //////    //////+ " ON sap_KEE_Resumen_ES_POR_MT_BTE_New.CUPS20 = SAP_PILOTO.CD_CUPS "
                            //////    + " WHERE fh_carga = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
                            //////    + " AND segmento IN ('MT','BTE')"
                            //////    //////+ " AND Piloto='" + Piloto + "'"
                            //////    + " GROUP BY   estado_global,   subestado, Incidencia_Facturacion, Incidencia_Medida"
                            //////    + " ORDER BY ESTADO_GLOBAL, subestado, Incidencia_Facturacion, Incidencia_Medida, PERIODO";
                            ///

                            strSql = "select A.estado_global,  A.subestado, A.Incidencia_Facturacion, A.Incidencia_Medida,"
                                   + " sum(A.Puno) as P1, sum(A.MayorPuno) as P2, sum(A.Tota) as Total"
                                   + " from"
                                   + " ("
                                   + " SELECT estado_global, subestado, Incidencia_Facturacion, Incidencia_Medida,"
                                   + " sum(case when periodo = '1P' then 1 else 0 end) AS Puno,"
                                   + " sum(case when periodo = '>1P' then 1 else 0 end) AS MayorPuno,"
                                   + " count(subestado) AS Tota"
                                   + " from facturacionb2b_owner.sap_kee_resumen_es_por"
                                   + " WHERE fh_carga = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
                                   + " AND segmento IN ('MT','BTE')"
                                   + " GROUP BY   estado_global, subestado, Incidencia_Facturacion, Incidencia_Medida, periodo"
                                   + " ) as A"
                                   + " group by A.estado_global,  A.subestado, A.Incidencia_Facturacion, A.Incidencia_Medida"
                                   + " ORDER BY A.estado_global, A.subestado, A.Incidencia_Facturacion, A.Incidencia_Medida ";


                        }
                        if (k == 3)
                            {
                                    //////strSql = "SELECT estado_global,  subestado, Incidencia_Facturacion, Incidencia_Medida, sum(CASE WHEN periodo = '1P' then 1 ELSE '' END) AS 'P1', sum(CASE WHEN periodo = '>1P' then 1 ELSE '' END) AS '>1P', COUNT(subestado) AS Total"
                                    //////    + " from sap_KEE_Resumen_ES_POR_MT_BTE_New "
                                    //////    //////+ " INNER JOIN SAP_PILOTO "
                                    //////    //////+ " ON sap_KEE_Resumen_ES_POR_MT_BTE_New.CUPS20 = SAP_PILOTO.CD_CUPS "
                                    //////    + " WHERE fh_carga = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
                                    //////    + " AND segmento IN ('BTN')"
                                    //////    //////+ " AND Piloto='" + Piloto + "'"
                                    //////    + " GROUP BY   estado_global,   subestado, Incidencia_Facturacion, Incidencia_Medida"
                                    //////    + " ORDER BY ESTADO_GLOBAL, subestado, Incidencia_Facturacion, Incidencia_Medida, PERIODO";

                                strSql = "select A.estado_global,  A.subestado, A.Incidencia_Facturacion, A.Incidencia_Medida,"
                                      + " sum(A.Puno) as P1, sum(A.MayorPuno) as P2, sum(A.Tota) as Total"
                                      + " from"
                                      + " ("
                                      + " SELECT estado_global, subestado, Incidencia_Facturacion, Incidencia_Medida,"
                                      + " sum(case when periodo = '1P' then 1 else 0 end) AS Puno,"
                                      + " sum(case when periodo = '>1P' then 1 else 0 end) AS MayorPuno,"
                                      + " count(subestado) AS Tota"
                                      + " from facturacionb2b_owner.sap_kee_resumen_es_por"
                                      + " WHERE fh_carga = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
                                      + " AND segmento IN ('BTN')"
                                      + " GROUP BY   estado_global, subestado, Incidencia_Facturacion, Incidencia_Medida, periodo"
                                      + " ) as A"
                                      + " group by A.estado_global,  A.subestado, A.Incidencia_Facturacion, A.Incidencia_Medida"
                                      + " ORDER BY A.estado_global, A.subestado, A.Incidencia_Facturacion, A.Incidencia_Medida ";
                            }

                        //////db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        //////command = new MySqlCommand(strSql, db.con);
                        //////r = command.ExecuteReader();
                        ///
                        dbRS = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD_FACTURACION);
                        commandRS = new OdbcCommand(strSql, dbRS.con);
                        rRS = commandRS.ExecuteReader();        

                        ExistenDatosCuadro = false;

                        while (rRS.Read())
                        {
                            ExistenDatosCuadro = true;
                            f++;
                            c = 2;

                            if (rRS["estado_global"] != System.DBNull.Value)
                            {
                                if (PintarEstado == "" || PintarEstado != rRS["estado_global"].ToString())
                                {
                                        ////workSheet.Cells[f, c].Value = r["estado_global"].ToString();
                                    if (PintarEstado != rRS["estado_global"].ToString() & PintarEstado != "")
                                    {
                                            //Paco 02/10/2024
                                            workSheet.Cells[CuadroInicio, 2].Value = PintarEstado;
                                            /////

                                            CuadroFin = f - 1;
                                            BordeCuadro = "B" + CuadroInicio.ToString() + ":F" + CuadroFin.ToString();

                                            ///Pinto Fila de Totales
                                            workSheet.Cells[f, 2].Value = " Total " + EtiquetaEstado;
                                            headerCells = workSheet.Cells[f, 2, f, 6];
                                            headerFont = headerCells.Style.Font;
                                            headerFont.Bold = true;
                                            workSheet.Cells[f, 6].Value = TotalParcial;
                                            TotalParcial = 0;
                                            f++;

                                            CuadroInicio = f;
                                            workSheet.Cells[BordeCuadro].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                            PintarSubestado = "";
                                    }
                                }

                                PintarEstado = rRS["estado_global"].ToString();
                                EtiquetaEstado = PintarEstado;
                                }
                                c++;

                                if (rRS["subestado"] != System.DBNull.Value)
                                {
                                    if (PintarSubestado == "" || PintarSubestado != rRS["subestado"].ToString())
                                    {
                                        workSheet.Cells[f, c].Value = rRS["subestado"].ToString();
                                    }
                                    PintarSubestado = rRS["subestado"].ToString();
                                }

                                c++;
                                if (rRS["Incidencia_Facturacion"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = rRS["Incidencia_Facturacion"].ToString();
                                }
                                c++;
                                if (rRS["Incidencia_Medida"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = rRS["Incidencia_Medida"].ToString();
                                }
                                c++;
                                ////////if (r["P1"] != System.DBNull.Value)
                                ////////    workSheet.Cells[f, c].Value = r["P1"].ToString();
                                ////////c++;
                                ////////if (r[">P1"] != System.DBNull.Value)
                                ////////    workSheet.Cells[f, c].Value = r[">P1"].ToString();
                                ////////c++;
                                if (rRS["Total"] != System.DBNull.Value)
                                {
                                    workSheet.Cells[f, c].Value = rRS["Total"];
                                    TotalParcial = TotalParcial + Convert.ToInt32(rRS["Total"]);
                                    TotalCuadro = TotalCuadro + Convert.ToInt32(rRS["Total"]);
                            }

                            c++;
                            allCells = workSheet.Cells[1, 1, f, c];
                            allCells.AutoFitColumns();
                        }

                        dbRS.CloseConnection();
                        strSql = "";

                        if (ExistenDatosCuadro == true)
                        {
                                //Pinto el recuadro final
                                //Paco 02/10/2024
                            workSheet.Cells[CuadroInicio, 2].Value = PintarEstado;
                                /////
                            CuadroFin = f;
                            BordeCuadro = "B" + CuadroInicio.ToString() + ":F" + CuadroFin.ToString();
                            workSheet.Cells[BordeCuadro].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                ///Pinto Fila de Estado
                            f++;
                            workSheet.Cells[f, 2].Value = " Total " + EtiquetaEstado;
                            workSheet.Cells[f, 6].Value = TotalParcial;
                            headerCells = workSheet.Cells[f, 2, f, 6];
                            headerFont = headerCells.Style.Font;
                            headerFont.Bold = true;
                            TotalParcial = 0;
                            f++;
                            workSheet.Cells[f, 1].Value = " Total " + Piloto;
                            workSheet.Cells[f, 6].Value = TotalCuadro;
                            headerCells = workSheet.Cells[f, 1, f, 6];
                            headerFont = headerCells.Style.Font;
                            headerFont.Bold = true;
                            TotalParcial = 0;
                            TotalCuadro = 0;

                            allCells = workSheet.Cells[1, 1, f, c];
                            allCells.AutoFitColumns();
                        }

                        //////} //// FIN for (int t = 1; t < 3; t++)
                    } //// for (int k = 1; k < 4; k++)
                    //// FIN  CREAMOS LAS  HOJAS RESUMEN EN TABLAS DINAMICAS *****************************************************************************

                    //Paco 16/02/2024
                    //Está grabado ya el detalle diario en las tablas de histórico, los puntos migrados los tenemos en las tablas t_ed_h_ps  y t_ed_h_ps_pt
                    //////select * from cont.t_ed_h_ps     where de_seg_mercado = 'SE'  and lg_migrado_sap = 'S';
                    //////SELECT * FROM t_ed_h_ps_pt where de_seg_mercado = 'GP'  and lg_migrado_sap = 'S'

                    //nombrecolumna = hojaActual.getCellByPosition(intColumna, 0).getColumns.getByIndex(0).getName;
                    ////nombrecolumna = Mid(hojaActual.Cells(1, intColumna).Address, 2, InStr(2, hojaActual.Cells(1, intColumna).Address, "$") - 2);
                    //////hojaActual.Cells["A1"].Value = "Esto funciona!";
                    //////var rangoDeCeldas = hojaActual.Cells["C1:C3"];
                    //////rangoDeCeldas = hojaActual.Cells["$C$1:$C$3"];

                    excelPackage.SaveAs(fileInfo);

                    //27/08/2024 Separar por columnas usando Aspose.Cells **************************
                    Workbook wb = new Workbook(ruta_salida_archivo);
   
                    TxtLoadOptions opts = new TxtLoadOptions();
                    opts.Separator = ';';

                    //Formateo de fechas
                    Style style = wb.CreateStyle();
                    style.Custom = "yyyy-mm-dd";

                    StyleFlag styleFlag = new StyleFlag();
                    styleFlag.NumberFormat = true;

                    for (int p = 1; p < 4; p++)
                    {
                        if (p == 1)
                        {
                            Worksheet sheet = wb.Worksheets["HISTORICO ES"];
                            sheet.Cells.TextToColumns(0, 0, TotalPruebaES, opts);
                            sheet.Cells.TextToColumns(0, 0, TotalPruebaES, opts);

                            Column column = sheet.Cells.Columns[3];
                            // Applying the style to the column
                            column.ApplyStyle(style, styleFlag);

                            column = sheet.Cells.Columns[4];
                            column.ApplyStyle(style, styleFlag);

                            for (int u = 26; u < 32; u++)
                            {
                                column = sheet.Cells.Columns[u];
                                column.ApplyStyle(style, styleFlag);
                            }
                            column = sheet.Cells.Columns[43];
                            column.ApplyStyle(style, styleFlag);
                            column = sheet.Cells.Columns[44];
                            column.ApplyStyle(style, styleFlag);
                            column = sheet.Cells.Columns[49];
                            column.ApplyStyle(style, styleFlag);
                            column = sheet.Cells.Columns[50];
                            column.ApplyStyle(style, styleFlag);
                            column = sheet.Cells.Columns[53];
                            column.ApplyStyle(style, styleFlag);

                            sheet.AutoFitColumns();

                        }
                        if (p == 2)
                        {
                            Worksheet sheet2 = wb.Worksheets["HISTORICO POR MT-BTE"];
                            sheet2.Cells.TextToColumns(0, 0, TotalPruebaMTBTE, opts);

                            Column column = sheet2.Cells.Columns[3];
                            // Applying the style to the column
                            column.ApplyStyle(style, styleFlag);

                            column = sheet2.Cells.Columns[4];
                            column.ApplyStyle(style, styleFlag);

                            for (int u = 26; u < 32; u++)
                            {
                                column = sheet2.Cells.Columns[u];
                                column.ApplyStyle(style, styleFlag);
                            }
;
                            column = sheet2.Cells.Columns[44];
                            column.ApplyStyle(style, styleFlag);
                            column = sheet2.Cells.Columns[45];
                            column.ApplyStyle(style, styleFlag);
                            column = sheet2.Cells.Columns[50];
                            column.ApplyStyle(style, styleFlag);
                            column = sheet2.Cells.Columns[51];
                            column.ApplyStyle(style, styleFlag);
                            column = sheet2.Cells.Columns[54];
                            column.ApplyStyle(style, styleFlag);
                            sheet2.AutoFitColumns();
                        }
                        if (p == 3)
                        {
                            Worksheet sheet3 = wb.Worksheets["HISTORICO POR BTN"];
                            sheet3.Cells.TextToColumns(0, 0, TotalPruebaBTN, opts);

                            Column column = sheet3.Cells.Columns[3];
                            // Applying the style to the column
                            column.ApplyStyle(style, styleFlag);
                            column = sheet3.Cells.Columns[4];
                            column.ApplyStyle(style, styleFlag);

                            for (int u = 26; u < 32; u++)
                            {
                                column = sheet3.Cells.Columns[u];
                                column.ApplyStyle(style, styleFlag);
                            }
                            column = sheet3.Cells.Columns[44];
                            column.ApplyStyle(style, styleFlag);
                            column = sheet3.Cells.Columns[45];
                            column.ApplyStyle(style, styleFlag);
                            column = sheet3.Cells.Columns[50];
                            column.ApplyStyle(style, styleFlag);
                            column = sheet3.Cells.Columns[51];
                            column.ApplyStyle(style, styleFlag);
                            column = sheet3.Cells.Columns[54];
                            column.ApplyStyle(style, styleFlag);
                            sheet3.AutoFitColumns();
                        }
                    }

                    ////////wb.worksheet("Evaluation Warning").IsVisible = false;
                    wb.Save(ruta_salida_archivo);

                    //Fin 27/08/2024 Separar por columnas usando Aspose.Cells ************************************

                    //29/08/2024 Abro con Interop.EXcel para ocultar la hoja que crea el expose ******************
                    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook librodetrabajo = excel.Workbooks.Open(ruta_salida_archivo);
                    ////Dimension = librodetrabajo.Worksheets.Count - 1;
                    //////librodetrabajo.Sheets["Evaluation Warning"].Remove();
                    //////librodetrabajo.Sheets["Evaluation Warning"].isVisible = false;
                    librodetrabajo.Sheets["Evaluation Warning"].Visible = Excel.XlSheetVisibility.xlSheetHidden;
                    ///librodetrabajo.Savechanged(ruta_salida_archivo);
                    librodetrabajo.Close(true);
                    excel.Quit();
                    //29/08/2024 Fin *****************************************************************************

                    /// MessageBox.Show("Fin");
                    //////EnvioCorreo_SAP_KEE(ruta_salida_archivo, fecha_informe, "Hoja ES (" + intConteoES.ToString() + "/" + intConteoESDetalle.ToString() + ") Procesados;Hoja MT-BTE (" + intConteoMTBTE.ToString() + "/" + intConteoMTBTEDetalle.ToString() + ") Procesados;Hoja BTN (" + intConteoBTN.ToString() + "/" + intConteoBTNDetalle.ToString() + ") Procesados",  SubestadosNoEncontrados.ToString(), FechaSapPendienteFacturar_fhenvio, FechaPdteWebB2B_fhact, FechaPdteWebB2B_fhInforme);

                    //////MessageBox.Show("Hoja ES (" + intConteoES.ToString() + "/" + intConteoESDetalle.ToString() + "') Procesados");
                    //////MessageBox.Show("Hoja MT-BTE (" + intConteoMTBTE.ToString() + "/" + intConteoMTBTEDetalle.ToString() + "') Procesados");
                    //////MessageBox.Show("BTN (" + intConteoBTN.ToString() + "/" + intConteoBTNDetalle.ToString() + "') Procesados");
                    //////MessageBox.Show("Hoja Subestado " + SubestadosNoEncontrados.ToString() + " no existe en la tabla paramétrica");
                    ///


                    //Paco 28/11/2024 separar fichero en dos, por un lado BTN y por otro el resto (hacer dos ficheros)

                    //ruta_salida_archivo = "C:\\Users\\fsanchezma\\Desktop\\InformePendienteOperaciones_20241127.xlsx";
                    Microsoft.Office.Interop.Excel.Application excel_ES_MT_BTE = new Microsoft.Office.Interop.Excel.Application();
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                    Microsoft.Office.Interop.Excel.Workbook librodetrabajo_ES_MT_BTE = excel_ES_MT_BTE.Workbooks.Open(ruta_salida_archivo);

                    excel_ES_MT_BTE.DisplayAlerts = false;
                    for (int u = librodetrabajo_ES_MT_BTE.Worksheets.Count; u >= 1; u--)
                    {
                        NombreHoja = librodetrabajo_ES_MT_BTE.Worksheets[u].Name;
                        if (NombreHoja.IndexOf("BTN") >= 0)
                        {

                            librodetrabajo_ES_MT_BTE.Sheets[u].Delete();

                        }
                    }

                    ruta_salida_archivo_BTE = ruta_salida_archivo.Replace("Operaciones", "Operaciones_ES_MT_BTE");
                    librodetrabajo_ES_MT_BTE.SaveCopyAs(ruta_salida_archivo_BTE);
                    librodetrabajo_ES_MT_BTE.SaveAs(false);
                    excel_ES_MT_BTE.DisplayAlerts = true;
                    librodetrabajo_ES_MT_BTE.Close(true);
                    excel_ES_MT_BTE.Quit();


                    //ruta_salida_archivo = "C:\\Users\\fsanchezma\\Desktop\\InformePendienteOperaciones_20241127.xlsx";
                    Microsoft.Office.Interop.Excel.Application excel_BTN = new Microsoft.Office.Interop.Excel.Application();
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                    Microsoft.Office.Interop.Excel.Workbook librodetrabajo_BTN = excel_BTN.Workbooks.Open(ruta_salida_archivo);

                    excel_BTN.DisplayAlerts = false;
                    for (int u = librodetrabajo_BTN.Worksheets.Count; u >= 1; u--)
                    {
                        NombreHoja = librodetrabajo_BTN.Worksheets[u].Name;
                        if (NombreHoja.IndexOf("BTN") < 0)
                        {

                            librodetrabajo_BTN.Sheets[u].Delete();

                        }
                    }

                    ruta_salida_archivo_BTN = ruta_salida_archivo.Replace("Operaciones", "Operaciones_BTN");
                    librodetrabajo_BTN.SaveCopyAs(ruta_salida_archivo_BTN);
                    librodetrabajo_BTN.SaveAs(false);
                    excel_BTN.DisplayAlerts = true;
                    librodetrabajo_BTN.Close(true);
                    excel_BTN.Quit();

                    //Fin Paco 28/11/2024 separar fichero en dos, por un lado BTN y por otro el resto (hacer dos ficheros)

                    if (param.GetValue("mail_enviar_mail_sap_kronos") == "S")
                    {
                        //Lanzamos ejecutable para subir a sharepoint el informe
                        utilidades.Fichero.EjecutaComando(param.GetValue("ruta_ejecutable_SubirSP_InformeSAPKEE"), null);

                        EnvioCorreo_SAP_KEE(ruta_salida_archivo_BTE, ruta_salida_archivo_BTN, fecha_informe, "Hoja ES (" + intConteoES.ToString() + "/" + intConteoESDetalle.ToString() + ") Procesados;Hoja MT-BTE (" + intConteoMTBTE.ToString() + "/" + intConteoMTBTEDetalle.ToString() + ") Procesados;Hoja BTN (" + intConteoBTN.ToString() + "/" + intConteoBTNDetalle.ToString() + ") Procesados", SubestadosNoEncontrados.ToString(), FechaSapPendienteFacturar_fhenvio, FechaPdteWebB2B_fhact, FechaPdteWebB2B_fhInforme);
                    }

                    //if (automatico && param.GetValue("mail_enviar_mail_psat_tam") == "S")
                    //{
                    //    ss_pp.Update_Fecha_Fin("Facturación", "Informe Pendiente KRONOS SAP BI", "Informe Pendiente KRONOS SAP BI");
                    //    EnvioCorreo_PdteWeb_BI(ruta_salida_archivo, fecha_informe);
                    //}

                    //if (automatico)
                    //    ss_pp.Update_Fecha_Fin("Facturación", "Informe Pendiente KRONOS SAP BI", "Informe Pendiente KRONOS SAP BI");

                }
                else if (automatico)
                {
                    ss_pp.Update_Comentario("Facturación", "Informe Pendiente KRONOS SAP BI", "Informe Pendiente KRONOS SAP BI",
                        "La fecha de actualización en BI es: " + UltimaActualizacionMySQL().Date.ToString("dd/MM/yyyy"));
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

            }
        }


        private static string ArreglaCadena(String t)
        {
            String salida;

            salida = t.Replace("\"", string.Empty); //string.Empty
            salida = salida.Replace("\\", string.Empty);
            salida = salida.Replace("\t", string.Empty);
            salida = salida.Replace("\n", string.Empty);
            salida = salida.Replace("\r", string.Empty);
            salida = salida.Replace("(\n|\r)", string.Empty);
            salida = salida.Replace("'", "´");
            salida = salida.Replace("Âª", string.Empty);
            salida = salida.Replace("_", "-");
            salida = salida.Trim();
            return salida;
        }

        private int FuncionDiasKEERecuperarDiasTabla(string Cups, string FechaDesde, string FechaHasta, string Estado)
        {
            Boolean blnLanzarDiasKEE;
            MySQLDB dbDiasKEE;
            MySqlDataReader DiasKEE;
            MySQLDB dbDiasKEEMax;
            MySqlDataReader DiasKEEMax;
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            string BuscaCadena = "";
            int Dias = 1;

            strSql = "select Dias as Numero from kee_Dias"
                   + " WHERE cups = '" + Cups + "'"
                   + " AND fh_desde = '" + FechaDesde + "'"
                   + " AND fh_hasta = '" + FechaHasta + "'"
                   + " AND estadoKEE = '" + Estado + "'"
                   + " AND Activo = true";

            dbDiasKEE = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, dbDiasKEE.con);
            DiasKEE = command.ExecuteReader();

            while (DiasKEE.Read())
            {
                Dias = Convert.ToInt32(DiasKEE["Numero"]) ;
            }

            dbDiasKEE.CloseConnection();
            return Dias;
        }

        private int FuncionDiasKEERecuperarDiasTablaGlobal(string Cups, string FechaDesde, string FechaHasta, string Estado)
        {
            Boolean blnLanzarDiasKEE;
            MySQLDB dbDiasKEE;
            MySqlDataReader DiasKEE;
            MySQLDB dbDiasKEEMax;
            MySqlDataReader DiasKEEMax;
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            string BuscaCadena = "";
            int Dias = 1;

            strSql = "select Dias as Numero from kee_Dias_Global"
                   + " WHERE cups = '" + Cups + "'"
                   + " AND fh_desde = '" + FechaDesde + "'"
                   + " AND fh_hasta = '" + FechaHasta + "'"
                   + " AND EstadoGlobal = '" + Estado + "'"
                   + " AND Activo = true";

            dbDiasKEE = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, dbDiasKEE.con);
            DiasKEE = command.ExecuteReader();

            while (DiasKEE.Read())
            {
                Dias = Convert.ToInt32(DiasKEE["Numero"]);
            }

            dbDiasKEE.CloseConnection();
            return Dias;
        }

        private int FuncionDiasKEE( string Cups, string FechaDesde, string FechaHasta, string Estado)
        {
            Boolean blnLanzarDiasKEE;
            MySQLDB dbDiasKEE;
            MySqlDataReader DiasKEE;
            MySQLDB dbDiasKEEMax;
            MySqlDataReader DiasKEEMax;
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            string BuscaCadena="";
            int Dias=1;

            strSql = "select Dias as Numero from kee_Dias"
                   + " WHERE cups = '" + Cups + "'"
                   + " AND fh_desde = '" + FechaDesde + "'"
                   + " AND fh_hasta = '" + FechaHasta + "'"
                   + " AND estadoKEE = '" + Estado + "'"
                   + " AND Activo = true";

             dbDiasKEE = new MySQLDB(MySQLDB.Esquemas.MED);
             command = new MySqlCommand(strSql, dbDiasKEE.con);
             DiasKEE = command.ExecuteReader();

             while (DiasKEE.Read())
             {
                if (Convert.ToInt32(DiasKEE["Numero"]) == 0)
                {
                    strSql = "INSERT INTO kee_Dias (CUPS, FH_DESDE, FH_HASTA, estadoKEE, Activo, Dias, FH_ACTUALIZACION) values ('"
                           + Cups + "','"
                           + FechaDesde + "','"
                           + FechaHasta + "','"
                           + Estado + "',true,1,'" + DateTime.Now.ToString("yyyy-MM-dd") + "')";
         

                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    Dias = 1;
                }
                else
                {

                    //Existe el registro, actualizo el numero de dias a dias+1 y PROCESADO_DIARIO='S'
                    strSql = "update kee_Dias  set"
                    + " DIAS = DIAS +1,"
                    + " FH_ACTUALIZACION='" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
                    + " WHERE cups = '" + Cups + "'"
                    + " AND fh_desde = '" + FechaDesde + "'"
                    + " AND fh_hasta = '" + FechaHasta + "'"
                    + " AND estadoKEE = '" + Estado + "'"
                    + " AND Activo = true";
        
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    Dias = Convert.ToInt32(DiasKEE["Numero"]) +1;
                }
              }

              dbDiasKEE.CloseConnection();
              return Dias;
        }


        private int FuncionDiasKEEGlobal(string Cups, string FechaDesde, string FechaHasta, string Estado)
        {
            Boolean blnLanzarDiasKEE;
            MySQLDB dbDiasKEE;
            MySqlDataReader DiasKEE;
            MySQLDB dbDiasKEEMax;
            MySqlDataReader DiasKEEMax;
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            string BuscaCadena = "";
            int Dias = 1;

            strSql = "select Dias as Numero from kee_Dias_Global"
                   + " WHERE cups = '" + Cups + "'"
                   + " AND fh_desde = '" + FechaDesde + "'"
                   + " AND fh_hasta = '" + FechaHasta + "'"
                   + " AND EstadoGlobal = '" + Estado + "'"
                   + " AND Activo = true";

            dbDiasKEE = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, dbDiasKEE.con);
            DiasKEE = command.ExecuteReader();

            while (DiasKEE.Read())
            {
                if (Convert.ToInt32(DiasKEE["Numero"]) == 0)
                {
                    strSql = "INSERT INTO kee_Dias_Global (CUPS, FH_DESDE, FH_HASTA, EstadoGlobal, Activo, Dias, FH_ACTUALIZACION) values ('"
                           + Cups + "','"
                           + FechaDesde + "','"
                           + FechaHasta + "','"
                           + Estado + "',true,1,'" + DateTime.Now.ToString("yyyy-MM-dd") + "')";


                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    Dias = 1;
                }
                else
                {

                    //Existe el registro, actualizo el numero de dias a dias+1 y PROCESADO_DIARIO='S'
                    strSql = "update kee_Dias_Global  set"
                    + " DIAS = DIAS +1,"
                    + " FH_ACTUALIZACION='" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
                    + " WHERE cups = '" + Cups + "'"
                    + " AND fh_desde = '" + FechaDesde + "'"
                    + " AND fh_hasta = '" + FechaHasta + "'"
                    + " AND  EstadoGlobal= '" + Estado + "'"
                    + " AND Activo = true";

                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    Dias = Convert.ToInt32(DiasKEE["Numero"]) + 1;
                }
            }

            dbDiasKEE.CloseConnection();
            return Dias;
        }

        private string EstadosAreas(string EstadoGlobal, string Cups, string Periodo, int HojaPintar, int NumeroDiasKEE, Boolean Incidencia)
        {

            MySQLDB dbReporte;
            MySqlCommand command;
            MySqlDataReader r;
            MySqlDataReader inci;
            MySqlDataReader reporte;
            string strSql = "";
            Boolean EstadosReporte;
            string cadena = "";
            string Auxiliar = "";

           
            if (HojaPintar < 3)
            {
                strSql = " select ESTADO_GLOBAL,ESTADO_GLOBAL_A_REPORTAR, SISTEMA  from estados_KEE_Param_EstadosReporte_New"
                       + " where Subestado='" + EstadoGlobal + "'"
                       + " and NivelTension <> 'BTN'";
            }
            else
            {
                strSql = " select ESTADO_GLOBAL,ESTADO_GLOBAL_A_REPORTAR, SISTEMA  from estados_KEE_Param_EstadosReporte_New"
                      + " where Subestado='" + EstadoGlobal + "'"
                      + " and NivelTension <> 'MT'";
            }
                                 

            dbReporte = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, dbReporte.con);
            reporte = command.ExecuteReader();

            EstadosReporte = false;
            while (reporte.Read())
            {
                EstadosReporte = true;

                Auxiliar = reporte["ESTADO_GLOBAL"].ToString();

                // para controlar si hay incidencia los que en el ESTADO_GLOBAL viene la etíqueta "SIN CODIGO"
                //////if (Auxiliar.Contains("SIN CÓDIGO"))
                //////{
                //////    if (Auxiliar.Contains("MEDIDA"))
                //////    {
                //////        if (Incidencia == true)
                //////        {
                //////            Auxiliar = "INCIDENCIAMEDIDA";
                //////        }

                //////        if (Auxiliar.Contains("FACTURACION"))
                //////        {
                //////            if (Incidencia == true)
                //////            {
                //////                Auxiliar = "INCIDENCIAFACTURACION";
                //////            }
                //////        }
                //////    }
                //////}

                switch (Auxiliar)
                {
                    case "FACTURACION":
                        // Si cruza,  los que estén en incidencia, los tiene que marcar como 02_INCIDENCIA FACTURACIÓN. Y en este caso, tendrán un código de incidencia. (02_INCIDENCIA FACTURACIÓN CON CÓDIGO - 02_INCIDENCIA )
                        // Si no cruza: MIRAR CAMPO DIAS_ESTADO  Si el periodo lleva en el mismo estado menos de 6 días "Estado transitorio, AVANZA" (05_AVANZA FACTURACIÓN - 05_AVANZA)
                        // Si el periodo lleva igual o más de 6 días "02_INCIDENCIA FACTURACIÓN".Y en este caso, no tendrán un código de incidencia. (021_INCIDENCIA FACTURACIÓN SIN CÓDIGO - 021_INCIDENCIA SIN COGIGO)

                        if (RelacionIncidenciasSAP.Cups(Cups + Periodo.ToString()))
                        {
                            cadena = "02_INCIDENCIA FACTURACION";
                            cadena= cadena + ";" + "02_INCIDENCIA";
                            cadena = cadena + ";" + "SISTEMAS FAC"; 
                        }
                        else
                        {
                            if (GetDiasEstado(Cups) >= 6)
                            {
                                cadena = "02_INCIDENCIA FACTURACION";
                                cadena = cadena + ";" + "02_INCIDENCIA";
                                cadena = cadena + ";" + "SISTEMAS FAC";
                            }
                            else
                            {
                                cadena = "05_AVANZA SISTEMAS";
                                cadena = cadena + ";" + "05_AVANZA";
                                cadena = cadena + ";" + "SISTEMAS FAC";
                            }
                        }
                        break;

                    case "02_INCIDENCIA_FACTURACION":

                        if (RelacionIncidenciasSAP.Cups(Cups + Periodo.ToString()))
                        {
                            cadena = "02_INCIDENCIA FACTURACION";
                            cadena = cadena + ";" + "02_INCIDENCIA";
                            cadena = cadena + ";" + "SISTEMAS FAC";
                        }
                        else
                        {
                            if (GetDiasEstado(Cups) >= 6)
                            {
                                cadena = "02_INCIDENCIA FACTURACION";
                                cadena = cadena + ";" + "02_INCIDENCIA";
                                cadena = cadena + ";" + "SISTEMAS FAC";
                            }
                            else
                            {
                                cadena = "05_AVANZA FACTURACION";
                                cadena = cadena + ";" + "05_AVANZA";
                                cadena = cadena + ";" + "FACTURACION";
                            }
                        }
                        break;

                    case "FACTURACION_AUX":

                        //Si el CUPS y PERIODO del pendiente cruza con el CUPS y PERIODO de la incidencia del fichero de incidencias ,entonces, automaticamente, el suministro se asigna a "02_INCIDENCIA FACTURACIÓN-02_INCIDENCIA
                        //No cruza : "04_PENDIENTE DE NEGOCIO - 04_PENDIENTE NEGOCIO FACTURACION"
                        if (RelacionIncidenciasSAP.Cups(Cups + Periodo.ToString()))
                        {
                            cadena = "02_INCIDENCIA FACTURACION";
                            cadena = cadena + ";" + "02_INCIDENCIA";
                            cadena = cadena + ";" + "SISTEMAS FAC";
                        }
                        else
                        {
                            cadena = "04_PENDIENTE NEGOCIO FACTURACION"; 
                            cadena = cadena + ";" + "04_PENDIENTE NEGOCIO";
                            cadena = cadena + ";" + "FACTURACION";
                        }
                        break;

                    case "FACTURACION_BIS":

                        //////Si el periodo lleva menos de 6 días ""Estado transitorio, AVANZA""
                        //////Si el periodo lleva igual o más de 6 días ""02_INCIDENCIA FACTURACIÓN"""

                        if (GetDiasEstado(Cups) >= 6)
                        {
                            cadena = "02_INCIDENCIA FACTURACION";
                            cadena = cadena + ";" + "02_INCIDENCIA";
                            cadena = cadena + ";" + "SISTEMAS FAC";
                        }
                        else
                        {
                            cadena = "05_AVANZA SISTEMAS";
                            cadena = cadena + ";" + "05_AVANZA";
                            cadena = cadena + ";" + "EN AVANCE";
                        }

                        break;

                    case "FACTURACION_DESCONOCIDO":

                        if (GetDiasEstado(Cups) >= 10)
                        {
                            cadena = "02_INCIDENCIA FACTURACION";
                            cadena = cadena + ";" + "02_INCIDENCIA";
                            cadena = cadena + ";" + "SISTEMAS FAC";
                        }
                        else {
                            cadena = "03_GAP FUNCIONAL";
                            cadena = cadena + ";" + "03_GAP FUNCIONAL";
                            cadena = cadena + ";" + "SISTEMAS FAC";
                        }

                        break;

                    case "MEDIDA":
                      
                        ///OLD
                        //////if (NumeroDiasKEE < 5)
                        //////{
                        //////    cadena = "05_AVANZA SISTEMAS";
                        //////    cadena = cadena + ";" + "05_AVANZA";
                        //////    cadena = cadena + ";" + "MEDIDA";
                        //////}
                        //////else
                        //////{
                        //////    if (Incidencia == true)
                        //////    {
                        //////        cadena = "02_INCIDENCIA_MEDIDA";
                        //////        cadena = cadena + ";" + "02_INCIDENCIA_MEDIDA";
                        //////        cadena = cadena + ";" + "SISTEMAS MED";
                        //////    }
                        //////    else
                        //////    {
                        //////        cadena = "021_INCIDENCIA_MEDIDA";
                        //////        cadena = cadena + ";" + "02_INCIDENCIA_MEDIDA";
                        //////        cadena = cadena + ";" + "SISTEMAS MED";
                        //////    } 
                        //////}
                        //////break;

                        ////Paco 29/11/2024, se cambia la lógica de evaluación
                        if (Incidencia == true)
                        {
                            cadena = "02_INCIDENCIA_MEDIDA";
                            cadena = cadena + ";" + "02_INCIDENCIA_MEDIDA";
                            cadena = cadena + ";" + "SISTEMAS MED";
                        }
                        else
                        {
                            if (NumeroDiasKEE < 5)
                            {
                                cadena = "05_AVANZA SISTEMAS";
                                cadena = cadena + ";" + "05_AVANZA";
                                cadena = cadena + ";" + "MEDIDA";
                            }
                           else
                            {
                                cadena = "021_INCIDENCIA_MEDIDA";
                                cadena = cadena + ";" + "02_INCIDENCIA_MEDIDA";
                                cadena = cadena + ";" + "SISTEMAS MED";
                            }       
                        }
                        break;
                        //// Fin Paco 29/11/2024

                    case "021_INCIDENCIA_MEDIDA":
                        cadena = "02_INCIDENCIA_MEDIDA";
                        cadena = cadena + ";" + "02_INCIDENCIA";
                        cadena = cadena + ";" + reporte["SISTEMA"].ToString();
                        break;

                    case "021_INCIDENCIA FACTURACION":

                        cadena = "02_INCIDENCIA FACTURACION";
                        cadena = cadena + ";" + "02_INCIDENCIA";
                        cadena = cadena + ";" + reporte["SISTEMA"].ToString();
                        break;

                    default:

                        cadena = reporte["ESTADO_GLOBAL"].ToString();
                        cadena = cadena + ";" + reporte["ESTADO_GLOBAL_A_REPORTAR"].ToString();
                        cadena = cadena + ";" + reporte["SISTEMA"].ToString();
                        break;

                }
            } // Fin  while (reporte.Read())

            if (EstadosReporte == false)
            {
                cadena = "";
                //////if (SubestadosNoEncontrados == "") {
                //////    SubestadosNoEncontrados = EstadoGlobal;
                //////}
                //////else {
                //////    SubestadosNoEncontrados = SubestadosNoEncontrados + ";" + EstadoGlobal;
                //////} 
            }
            dbReporte.CloseConnection();
            return cadena;
        }

        private string Reincidente(string areaincidencia, string arearesponsable)
        {
            string Dato = "";

            if (arearesponsable != "" && arearesponsable != null) //He cruzado con KEE
            {
                if (areaincidencia.ToUpper() == arearesponsable.ToUpper())
                {
                    if (areaincidencia == "Contratacion")
                        Dato = "01.B09 Incidencia_Contratacion";
                    if (areaincidencia == "Medida")
                        Dato = "01.B10 Incidencia_Medida";
                    if (areaincidencia == "Facturacion")
                        Dato = "01.B11 Incidencia_Facturacion";
                }
            }
            else
            { // No cruza con KEE
                if (areaincidencia == "Facturacion")
                {
                    Dato = "01.B11 Incidencia_Facturacion"; 
                }
                ////Paco 18/12/2024
                if (areaincidencia == "Medida")
                {
                    Dato = "01.B10 Incidencia_Medida";
                }

                //fin Paco 18/12/2024
            } // Fin  if (subestadosKronos.area_responsable != "")

            return Dato;

        }

        private int Total_Pendiente(bool agora, DateTime fecha, List<string> lista_empresas, List<string> lista_segmentos, string estado, string subestado)
        {
            int total = 0;

            List<EndesaEntity.medida.Pendiente> o;

            if (lista_segmentos == null)
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int i = 0; i < o.Count; i++)
                        {
                            /*
                            if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora &&
                                (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null) ))
                                total = total + 1;
                            */

                            if (o[i].subestado_SAP is null) {
                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora && (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                    total = total + 1;
                            }
                            else {
                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora && (o[i].cod_estado == estado && (o[i].cod_subestado_SAP == subestado || subestado == null)))
                                    total = total + 1;
                            }


                        }
                }
            }
            else
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int z = 0; z < lista_segmentos.Count; z++)
                            for (int i = 0; i < o.Count; i++)
                            {
                                /*
                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z]
                                    && o[i].agora == agora &&
                                    (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                    total = total + 1;
                                */

                                if (o[i].subestado_SAP == null)
                                {
                                    if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z]  && o[i].agora == agora &&
                                    (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                        total = total + 1;
                                }
                                else
                                {

                                   if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z]  && o[i].agora == agora &&
                                   (o[i].cod_estado == estado && (o[i].cod_subestado_SAP == subestado || subestado == null)))
                                        total = total + 1;

                                }

                            }
                }
            }

            return total;

        }

        private int Total_Pendiente_25(bool agora, DateTime fecha, List<string> lista_empresas, List<string> lista_segmentos, string estado, string subestado, string Etiqueta)
        {
            int total = 0;

            List<EndesaEntity.medida.Pendiente> o;

            if (lista_segmentos == null)
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int i = 0; i < o.Count; i++)
                        {
                            if (o[i].cod_subestado_SAP == "01.C25 Pendiente Medida KEE - En ciclo de Medida") {
                                System.Diagnostics.Debug.WriteLine("");
                            }
                            //System.Diagnostics.Debug.WriteLine(c.cups20 +

                            if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora && o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null) && o[i].subestado_SAP == subestado + " " + Etiqueta)
                                    total = total + 1;

                        }
                }
            }
            else
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int z = 0; z < lista_segmentos.Count; z++)
                            for (int i = 0; i < o.Count; i++)
                            {
                                /*
                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z]
                                    && o[i].agora == agora &&
                                    (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                    total = total + 1;
                                */

                                if (o[i].cod_subestado_SAP == "01.C25")
                                {
                                    System.Diagnostics.Debug.WriteLine("");
                                 }

                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z] && o[i].agora == agora &&
                                    o[i].cod_estado == estado && (o[i].cod_subestado_SAP == subestado || subestado == null) && o[i].subestado_SAP == subestado + " " + Etiqueta)
                                        total = total + 1;

                             

                            }
                }
            }

            return total;

        }

        private int Total_PendienteUno(bool agora, DateTime fecha, List<string> lista_empresas, List<string> lista_segmentos, string estado, string subestado, string SAP )
        {
            int total = 0;

            List<EndesaEntity.medida.Pendiente> o;

            if (lista_segmentos == null)
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int i = 0; i < o.Count; i++)
                        {
                            /*
                            if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora &&
                                (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null) ))
                                total = total + 1;
                            */

                            //if (dia.ToString("yyyy-MM-dd") == o[i].fecha_informe.ToString("yyyy-MM-dd")){ 

                                if (SAP == "N") {
                                    if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora && (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)) && (o[i].cod_subestado_SAP==null))
                                        total = total + 1;
                                }
                                else
                                {
                                    if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora && (o[i].cod_estado == estado && (o[i].cod_subestado_SAP == subestado || subestado == null)) )
                                        total = total + 1;
                                }

                            //}
                        }
                }
            }
            else
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int z = 0; z < lista_segmentos.Count; z++)
                            for (int i = 0; i < o.Count; i++)
                            {
                                /*
                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z]
                                    && o[i].agora == agora &&
                                    (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                    total = total + 1;
                                */

                               /* if (o[i].subestado_SAP == null)
                                {
                                    if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z] && o[i].agora == agora &&
                                    (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                        total = total + 1;
                                }
                                else
                                {

                                    if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z] && o[i].agora == agora &&
                                    (o[i].cod_estado == estado && (o[i].cod_subestado_SAP == subestado || subestado == null)))
                                        total = total + 1;

                                }
                                */

                                    if (SAP == "N")
                                    {
                                        if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z] && o[i].agora == agora &&
                                        (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                            total = total + 1;
                                    }

                                    // if (o[i].subestado_SAP is null)
                                    //{
                                    //   if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora && (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                    //       total = total + 1;
                                    //}
                                    else
                                    {
                                        if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z] && o[i].agora == agora &&
                                    (o[i].cod_estado == estado && (o[i].cod_subestado_SAP == subestado || subestado == null)))
                                            total = total + 1;
                                    }


                            }
                }
            }

            return total;

        }

        private double Total_Pendiente_TAM(bool agora, DateTime fecha, List<string> lista_empresas, List<string> lista_segmentos, string estado, string subestado)
        {
            double total = 0;

            List<EndesaEntity.medida.Pendiente> o;

            if (lista_segmentos == null)
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int i = 0; i < o.Count; i++)
                        {

                            /*
                            if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora &&
                                (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                total = total + o[i].tam;
                            */

                            if (o[i].subestado_SAP == null)
                            {
                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora && (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                    total = total + o[i].tam;
                            }
                            else
                            {
                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora && (o[i].cod_estado == estado && (o[i].cod_subestado_SAP == subestado || subestado == null)))
                                    total = total + o[i].tam;
                            }

                        }


                }
            }
            else
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int z = 0; z < lista_segmentos.Count; z++)
                            for (int i = 0; i < o.Count; i++)
                            {
                                /*
                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z]
                                        && o[i].agora == agora &&
                                    (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                    total = total + o[i].tam;
                                */

                                if (o[i].subestado_SAP == null)
                                {
                                    if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z] && o[i].agora == agora &&
                                    (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                        total = total + o[i].tam;
                                }
                                else
                                {

                                    if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z] && o[i].agora == agora &&
                                    (o[i].cod_estado == estado && (o[i].cod_subestado_SAP == subestado || subestado == null)))
                                        total = total + o[i].tam;

                                }

                            }

                }
            }

            return total;

        }

        private double Total_Pendiente_TAM_25(bool agora, DateTime fecha, List<string> lista_empresas, List<string> lista_segmentos, string estado, string subestado, string Etiqueta)
        {
            double total = 0;

            List<EndesaEntity.medida.Pendiente> o;

            if (lista_segmentos == null)
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int i = 0; i < o.Count; i++)
                        {
                         


                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora && o[i].cod_estado == estado && o[i].cod_subestado_SAP == subestado && o[i].subestado_SAP == subestado + " " + Etiqueta)
                                    total = total + o[i].tam;
                        
                        }
                }
            }
            else
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int z = 0; z < lista_segmentos.Count; z++)
                            for (int i = 0; i < o.Count; i++)
                            {
                                /*
                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z]
                                        && o[i].agora == agora &&
                                    (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                    total = total + o[i].tam;
                                */
                                if (o[i].cod_subestado_SAP == "01.C25")
                                {
                                    System.Diagnostics.Debug.WriteLine("");
                                }


                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z] && o[i].agora == agora &&
                                    o[i].cod_estado == estado && o[i].cod_subestado_SAP == subestado && o[i].subestado_SAP == subestado + " " + Etiqueta)
                                        total = total + o[i].tam;

                              

                            }

                }
            }

            return total;

        }

        private string DetalleExcel(DateTime fecha_informe, List<string> lista_empresas, List<string> lista_tension)
        {
            string strSql = "";
            bool firstOnly = true;


            strSql = "SELECT ps.cd_empr, p.cd_segmento_ptg, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, p.cd_cups as cups20, p.id_instalacion, ps.cd_tarifa_c,"
                + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre, p.lg_multimedida, p.fh_desde, p.fh_hasta, p.cd_subestado,"
                + " concat(p.cd_estado,' ',de.de_estado) as de_estado, concat(p.cd_subestado,' ',if (ds.de_subestado is null,'', ds.de_subestado)) as de_subestado, p.fh_periodo as mes, IF(s.CUPS20 IS NULL, p.agora, 'S') AS agora, p.TAM , ps.fh_prev_fin_crto,  ps.fh_baja_crto, p.id_crto_ext,p.fec_act";
            //07/11/2024
            if (lista_empresas[0].Contains("PT"))
                strSql += " ,case when p.cd_segmento_ptg is null then segmento ELSE p.cd_segmento_ptg END AS cd_segmento_ptg_Con_compor";
            // Fin  07/11/2024
            strSql += " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p";

            if (lista_empresas[0].Contains("PT"))
                strSql += " LEFT OUTER JOIN cont.t_ed_h_ps_pt ps ON";
            else
                strSql += " LEFT OUTER JOIN cont.t_ed_h_ps ps ON";

            strSql += " ps.cups20 = p.cd_cups"
                + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                + " de.cd_estado = p.cd_estado"
                + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                + " ds.cd_subestado = p.cd_subestado"

            // 06/02/2025 añadir sofisticados agora
            +" LEFT OUTER JOIN fact.cm_sofisticados s on s.CUPS20 = p.cd_cups";
            //Fin 06/02/2025

            //07/11/2024
            if (lista_empresas[0].Contains("PT"))
            {
                ///07-11-2024 Para recuperar los segmentos de mercado nulo
                strSql += " LEFT JOIN fact.t_ed_h_sap_pendiente_facturar_segmentos_compor"
                + " ON p.cd_cups = t_ed_h_sap_pendiente_facturar_segmentos_compor.cd_cups ";
                /// Fin 07/11/2024 Para recuperar los segmentos de mercado nulo
            }
            // Fin  07/11/2024

            strSql += " where p.fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "'";
            //////+ " and p.cd_cups in ('PT0002000111610262EM')";
            ///
            //////+ " and p.cd_cups in ( 'PT0002000023467508BJ','PT0002000004930716FW')";
            ///
            //07/11/2024
            if (lista_empresas[0].Contains("PT"))
            {
                strSql += "  AND (p.fh_desde >= t_ed_h_sap_pendiente_facturar_segmentos_compor.fdesde "
                + " AND t_ed_h_sap_pendiente_facturar_segmentos_compor.fdesde <= p.fh_hasta) ";
            }

            // Fin  07/11/2024
            //Hay que quitar (porque así se decidió en la Negociación de este reporte) que no pueden aparecer registros con el mes pendiente = al mes en curso !!!!!!!!!!!
            strSql += " and p.fh_periodo <> '" + DateTime.Now.ToString("yyyyMM") + "'"
                //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                +" and p.cd_empr_titular in (";

                foreach (string p in lista_empresas)
                {
                    if (firstOnly)
                    {
                        strSql += "'" + p + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + p + "'";

                }

                strSql += ")";

            //////if (lista_tension != null)
            //////{
            //////firstOnly = true;
            //////strSql += " and p.cd_segmento_ptg in (";
            //////foreach (string p in lista_tension)
            //////{
            //////    if (firstOnly)
            //////    {
            //////        strSql += "'" + p + "'";
            //////        firstOnly = false;
            //////        Segmentos= "'" + p + "'";
            //////    }
            //////    else
            //////    {
            //////        strSql += ",'" + p + "'";
            //////        Segmentos= Segmentos + ",'" + p + "'";
            //////    }


            //////}
            //////strSql += ")";
            ///

            //07/11/2024
            if (lista_tension != null)
            {
                firstOnly = true;
                strSql += " and (p.cd_segmento_ptg in (";
                foreach (string p in lista_tension)
                {
                    if (firstOnly)
                    {
                        strSql += "'" + p + "'";
                        firstOnly = false;
                    }
                    else
                    {
                        strSql += ",'" + p + "'";

                    }
                }
                strSql += ") or p.cd_segmento_ptg  is null ) ";
            }

            strSql += " GROUP BY ps.cd_empr, p.cd_segmento_ptg, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli, ps.fh_alta_crto, ps.fh_inicio_vers_crto, p.cd_cups , p.id_instalacion,"
                 + " ps.cd_tarifa_c, ps.cd_crto_comercial, ps.de_empr_distdora_nombre, p.lg_multimedida, p.fh_desde, p.fh_hasta, p.cd_subestado, "
                 + " concat(p.cd_estado, ' ', de.de_estado) , concat(p.cd_subestado, ' ',if (ds.de_subestado is null,'', ds.de_subestado)) , "
                 + " p.fh_periodo , IF(s.CUPS20 IS NULL, p.agora, 'S') , p.TAM , ps.fh_prev_fin_crto,  ps.fh_baja_crto, p.id_crto_ext,p.fec_act ";

            if (lista_empresas[0].Contains("PT"))
                strSql += " ,case when p.cd_segmento_ptg is null then segmento ELSE p.cd_segmento_ptg END";
                   
            //Fin 07/11/2024
            return strSql;

        }
        private void EnvioCorreo_PdteWeb_BI(string archivo, DateTime fecha_informe)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                string from = param.GetValue("mail_from_psat_tam");
                string to = param.GetValue("mail_to_psat_tam");
                string cc = param.GetValue("mail_cc_psat_tam");
                string subject = param.GetValue("mail_subject_psat_tam") + " a " + fecha_informe.ToString("dd/MM/yyyy");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Se adjunta el archivo ").Append(fileInfo.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");

                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (param.GetValue("mail_enviar_mail_psat_tam") == "S")
                    mes.SendMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);

                else
                    mes.SaveMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("mail_from"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }

        private void EnvioCorreo_SAP_KEE(string archivo_bte, string archivo_btn, DateTime fecha_informe, string Procesados, string Estados, string FechaSapPendienteFacturar, string FechaPdteWebB2B, string FechaPdteWebB2B_fhInforme)
        {
            FileInfo fileInfo_BTE = new FileInfo(archivo_bte);
            FileInfo fileInfo_BTN = new FileInfo(archivo_btn);
            StringBuilder textBody = new StringBuilder();

            try
            {
                string from = param.GetValue("mail_from_sap_kronos");
                string to = param.GetValue("mail_to_sap_kronos");
                string cc = param.GetValue("mail_cc_sap_kronos");
                string subject = param.GetValue("mail_subject_sap_kronos") + " - " + fecha_informe.ToString("dd/MM/yyyy");

                //to = "francisco.sanchezmartin@enel.com";
                //cc = "gbenavides@minsait.com";
                //cc = "eloy.fernandezg@enel.com;adriana.lopezl@enel.com;angel.diazm@enel.com;francisco.sanchezmartin@enel.com;scasati.rivera@servinform.es;lmauricio.carrasco@servinform.es;brodriguez.carmona@servinform.es;ajtinoco.sanchez@servinform.es;gbenavides@minsait.com;fenix_opm_facturacion@servinform.es;anamaria.alonso@enel.com";
                //subject = "Informe Pendiente Operaciones " + "-" + fecha_informe.ToString("dd/MM/yyyy");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Subido el fichero a Enel SPA\\Seguimiento Pendiente Fenix\\General ").Append(fileInfo_BTE.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Disponible en ");
                textBody.Append("<a href=\"https://enelcom.sharepoint.com/sites/SEGUIMIENTOPENDIENTEFENIX/Shared%20Documents/General/").Append(fileInfo_BTE.Name).Append("\">Informe Pendiente Operaciones ES_MT_BTE</a>");
                textBody.Append(System.Environment.NewLine);

                textBody.Append("  Subido el fichero a Enel SPA\\Seguimiento Pendiente Fenix\\General ").Append(fileInfo_BTN.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Disponible en ");
                textBody.Append("<a href=\"https://enelcom.sharepoint.com/sites/SEGUIMIENTOPENDIENTEFENIX/Shared%20Documents/General/").Append(fileInfo_BTN.Name).Append("\">Informe Pendiente Operaciones BTN</a>");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                if (Procesados != "")
                {
                    string[] cadena = Procesados.Split(';');
                    textBody.Append("  - " + cadena[0]);
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("  - " + cadena[1]);
                    textBody.Append(System.Environment.NewLine);
                    textBody.Append("  - " + cadena[2]);
                    textBody.Append(System.Environment.NewLine);
                }

                if (Estados != "")
                {
                    string[] cadena = Estados.Split(';');
                    foreach (string a in cadena)
                    {
                        textBody.Append("  - " + a);
                        textBody.Append(System.Environment.NewLine);
                    }
                }

                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("- fh_envio de la tabla t_ed_h_sap_pendiente_facturar: " + FechaSapPendienteFacturar);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("- fec_act de la tabla t_ed_h_pdtweb_pm_b2b: " + FechaPdteWebB2B);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("- fecha_informe de la tabla t_ed_h_pdtweb_pm_b2b: " + FechaPdteWebB2B_fhInforme);
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Atencion!! Es necesario tener actualizado el Excel de INCIDENCIAS_FACTURACION para que el resultado sea el esperado.");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Por favor, informar OK/KO del informe para poder avanzar en el desarrollo del aplicativo");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);

                textBody.Append("Un saludo.");

                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                //////if (param.GetValue("mail_enviar_mail_psat_tam") == "S")
                // GUS: Modificamos para no enviar adjunto - 18-09-2024
                //mes.SendMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);
                mes.SendMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml_href(textBody.ToString()), "");
                //////else
                //////    mes.SaveMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);

                //////ficheroLog.Add("Correo enviado desde: " + param.GetValue("mail_from"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }

        //////private static void ShowSBInfo(StringBuilder sb)

        //////{
        //////    foreach (var prop in sb.GetType().GetProperties()) {
        //////        if (prop.GetIndexParameters().Length == 0)
        //////           Console.Write ("{0}: {1:N0}     ", prop.Name, prop.GetValue(sb));
        //////    }
        //////}
        private DateTime CalculaFechaDesdeinforme()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime date = DateTime.Now;

            strSql = "SELECT c.fh_envio"
                + " FROM t_ed_h_sap_pendiente_facturar_agrupado c"
                + " GROUP BY c.fh_envio desc"
                + " LIMIT 5";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["fh_envio"] != System.DBNull.Value)
                    date = Convert.ToDateTime(r["fh_envio"]);
            }
            db.CloseConnection();

            return date;

        }

        private DateTime CalculaFechaDesdeinformeMaxPendiente ()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime date = DateTime.Now;

            strSql = "SELECT c.fh_envio"
                + " FROM t_ed_h_sap_pendiente_facturar_agrupado c"
                + " GROUP BY c.fh_envio order by c.fh_envio desc"
                + " LIMIT 5";



            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["fh_envio"] != System.DBNull.Value)
                    date = Convert.ToDateTime(r["fh_envio"]);
                break;
            }
            db.CloseConnection();

            return date;

        }


        private DateTime CalculaFechaDesdeinformeMaxPendienteDesdeFechaEnvio()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime date = DateTime.Now;

            //////strSql = "SELECT c.fh_envio"
            //////    + " FROM t_ed_h_sap_pendiente_facturar_agrupado c"
            //////    + " GROUP BY c.fh_envio desc"
            //////    + " LIMIT 5";

            //11-04-2025 
            strSql = "SELECT A.fh_envio FROM( "
             + " SELECT c.fh_envio FROM t_ed_h_sap_pendiente_facturar_agrupado c GROUP BY c.fh_envio order by c.fh_envio desc LIMIT 5 "
             + " ) AS A "
             + " ORDER BY A.fh_envio asc "
             + " LIMIT 1 ";
            ///

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["fh_envio"] != System.DBNull.Value)
                    date = Convert.ToDateTime(r["fh_envio"]);
                break;
            }
            db.CloseConnection();

            return date;

        }

        private DateTime CalculaFechaMaxPdteWebB2B()
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";
            DateTime date = DateTime.Now;

            strSql = "SELECT max(fec_act) as fecha"
                + " FROM ed_owner.t_ed_h_pdtweb_pm_b2b";

            db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
            command = new OdbcCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())

            {
                if (r["fecha"] != System.DBNull.Value)
                    date = Convert.ToDateTime(r["fecha"]);
                break;
            }
            db.CloseConnection();

            return date;

        }

        private DateTime CalculaFechaMaxPdteWebB2B_fhinforme(DateTime fechaSAP)
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";
            DateTime date = DateTime.Now;


            strSql = "SELECT max(fecha_informe) as fecha"
                + " FROM ed_owner.t_ed_h_pdtweb_pm_b2b"
                + " where fecha_informe<'" + fechaSAP.Date.ToString("yyyy-MM-dd") + "'";
          
            db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
            command = new OdbcCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())

            {
                if (r["fecha"] != System.DBNull.Value)
                    date = Convert.ToDateTime(r["fecha"]);
                break;
            }
            db.CloseConnection();

            return date;

        }

        private DateTime CalculaFechaDetalle()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime date = DateTime.Now;

            strSql = "SELECT max(c.fh_envio) as max_fecha"
                + " FROM t_ed_h_sap_pendiente_facturar_agrupado c";


            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["max_fecha"] != System.DBNull.Value)
                    date = Convert.ToDateTime(r["max_fecha"]);
            }
            db.CloseConnection();

            return date;

        }

        private Dictionary<string, DateTime> CargaDiasEstado()
        {
            Dictionary<string, DateTime> d = new Dictionary<string, DateTime>();

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string cups = "";
            DateTime primera_fecha = DateTime.Now;

            strSql = "SELECT p.cd_cups, p.fh_periodo, MIN(p.fh_envio) AS primera_fecha, p.cd_subestado"
                + " FROM t_ed_h_sap_pendiente_facturar_agrupado p"
                + " WHERE p.fh_periodo >= " + DateTime.Now.AddYears(-1).ToString("yyyyMM")
                + " GROUP BY p.cd_cups, p.fh_periodo, p.cd_subestado"
                + " ORDER BY p.fh_envio DESC";


            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                cups = r["cd_cups"].ToString();
                primera_fecha = Convert.ToDateTime(r["primera_fecha"]);

                DateTime o;
                if (!d.TryGetValue(cups, out o))
                    d.Add(cups, primera_fecha);

            }
            db.CloseConnection();
            return d;

        }

        private int GetDiasEstado(string cups)
        {
            int dias = 1;
            DateTime o;
            if (dic_dias_estado.TryGetValue(cups, out o))
            {
                dias = (DateTime.Now.Date - o.Date).Days + 1;
            }



            return dias;
        }
    }
}