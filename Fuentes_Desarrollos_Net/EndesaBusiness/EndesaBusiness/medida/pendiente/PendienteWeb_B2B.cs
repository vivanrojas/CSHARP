using EndesaBusiness.servidores;
using EndesaEntity.medida;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace EndesaBusiness.medida.pendiente
{
    public class PendienteWeb_B2B : PendienteMedida_B2B
    {
        logs.Log ficheroLog;
        Dictionary<string, List<PendienteMedida_B2B>> dic;
        public bool existe { get; set; }

        public Dictionary<string, PendienteMedida_B2B> relac { get; set; }

        public PendienteWeb_B2B()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "PendienteMedida_B2B");



            //dic = Carga(MaxFechaInforme()); //Saca el maximo de t_ed_h_pdtweb_pm_b2b
            // Hay problemas de sincronización, KEE sale todos los días, en cambio SAP no se ejecuta fines de semana y ¿Festivos?
            // A la hora de ir a buscar a KEE, buscamos la fecha maxima de SAP para sincronizar los dos
            dic = Carga(CalculaFechaDesdeinformeMaxPendiente());


        }

        public PendienteWeb_B2B(int uno)
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "PendienteMedida_B2B");

            if (uno == 1) //Fecha Baja SAP
            {
                dic = CargaFechaBajaSAP();
            }
            else {
                dic = RelacionIncidenciasCUPS();
            }

        }


        public PendienteWeb_B2B(string strBTN)
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "PendienteMedida_B2B");



            //dic = Carga(MaxFechaInforme()); //Saca el maximo de t_ed_h_pdtweb_pm_b2b
            // Hay problemas de sincronización, KEE sale todos los días, en cambio SAP no se ejecuta fines de semana y ¿Festivos?
            // A la hora de ir a buscar a KEE, buscamos la fecha maxima de SAP para sincronizar los dos
            dic = CargaBTN(CalculaFechaDesdeinformeMaxPendiente());


        }

        //////public PendienteWeb_B2B(bool RelacionIncidencias)
        //////{

        //////    //dic = Carga(MaxFechaInforme()); //Saca el maximo de t_ed_h_pdtweb_pm_b2b
        //////    // Hay problemas de sincronización, KEE sale todos los días, en cambio SAP no se ejecuta fines de semana y ¿Festivos?
        //////    // A la hora de ir a buscar a KEE, buscamos la fecha maxima de SAP para sincronizar los dos
        //////    dic = CargaRelacionIncidencias();

        //////}

        //////private Dictionary<string, List<PendienteMedida_B2B>> CargaRelacionIncidencias()
        //////{
        //////    MySQLDB db;
        //////    MySqlCommand command;
        //////    MySqlDataReader r;
        //////    string strSql = "";

        //////    ////fecha_ControlCargas = MaxFechaInformeGestDiarioB2B();
          
        //////    Dictionary<string, PendienteMedida_B2B> relac = new Dictionary<string, PendienteMedida_B2B>();

        //////    try
        //////    {

        //////        strSql = "select area, cups, mes_pendiente, incidencia, estado_incidencia, fecha_apertura, prioridad_negocio, titulo, e_s_estado, Reincidente from Relacion_INC_CUPS order by Fecha_apertura asc";

        //////        db = new MySQLDB(MySQLDB.Esquemas.MED);
        //////        command = new MySqlCommand(strSql, db.con);
        //////        r = command.ExecuteReader();
        //////        while (r.Read())
        //////        {

        //////            PendienteMedida_B2B c = new PendienteMedida_B2B();
        //////            c.area = r["area"].ToString();
        //////            c.cups = r["cups"].ToString();
        //////            c.incidencia = r["incidencia"].ToString();
        //////            c.estado_incidencia = r["estado_incidencia"].ToString();
        //////            c.prioridad_negocio = r["prioridad_negocio"].ToString();
        //////            c.titulo = r["titulo"].ToString();
        //////            c.e_s_estado = r["e_s_estado"].ToString();
        //////            if (r["fecha_apertura"] != System.DBNull.Value)
        //////            {
        //////                c.fecha_apertura = Convert.ToDateTime(r["fecha_apertura"]);
        //////            }
        //////            c.Reincidente = r["Reincidente"].ToString();
        //////            c.mes_pendiente = r["mes_pendiente"].ToString();

        //////            //////List<PendienteMedida_B2B> t;
        //////            //////if (!d.TryGetValue(c.cups + c.mes_pendiente + ";" + c.incidencia  + c.area, out t))
        //////            //////{
        //////            //////    t = new List<PendienteMedida_B2B>();
        //////            //////    t.Add(c);
        //////            //////    d.Add(c.cups + c.mes_pendiente + ";" + c.incidencia + c.area, t);
        //////            //////}
        //////            //////else
        //////            //////    t.Add(c);

        //////            relac.Add(c.cups + c.mes_pendiente, c);
        //////        }
        //////        db.CloseConnection();

        //////        return relac;
        //////    }
        //////    catch (Exception ex)
        //////    {
        //////        ficheroLog.addError("Carga: " + ex.Message);
        //////        return null;
        //////    }
        //////}

        public List<string> GetIncidencia(string cups, string Mes)
        {
            List<string> lista = new List<string>();

            List<PendienteMedida_B2B> o;
            if (dic.TryGetValue(cups + Mes + "%", out o))
            {
                lista = o.FindAll(z => z.cups == cups & z.mes_pendiente == Mes).Select(z => Convert.ToString(z.multipunto)).ToList();
            }
            return lista;
        }

        //////private DateTime MaxFechaInformeGestDiarioB2B_BTN()
        //////{
        //////    // Buscamos la última fecha de publicación del informe

        //////    servidores.RedShiftServer db;
        //////    OdbcCommand command;
        //////    OdbcDataReader r;
        //////    string strSql = "";
        //////    DateTime fecha_max = new DateTime();

        //////    try
        //////    {
        //////        //strSql = "select max(fh_ejecucion) as max_fecha from ed_owner.T_ED_H_GEST_DIAR_PS_B2B";

        //////        strSql = "select max(fec_act) as max_fecha from ed_owner.t_ed_h_sap_pendiente_facturar";

        //////        // ed_owner.t_ed_h_sap_pendiente_facturar 
        //////        ficheroLog.Add(strSql);
        //////        db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
        //////        command = new OdbcCommand(strSql, db.con);
        //////        r = command.ExecuteReader();
        //////        while (r.Read())
        //////        {
        //////            fecha_max = Convert.ToDateTime(r["max_fecha"]);
        //////        }
        //////        db.CloseConnection();

        //////        return fecha_max;
        //////    }
        //////    catch (Exception ex)
        //////    {
        //////        ficheroLog.addError("MaxFechaInforme: " + ex.Message);
        //////        return DateTime.Now;
        //////    }
        //////}
        ///

        private Dictionary<string, List<PendienteMedida_B2B>> CargaBTN(DateTime fecha_max)
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";

           ////fecha_ControlCargas = MaxFechaInformeGestDiarioB2B();
            Dictionary<string, List<PendienteMedida_B2B>> d = new Dictionary<string, List<PendienteMedida_B2B>>();

            try
            {
                //Paco- añadir fecha baja contrato KEE, hay que hacer left join con T_ED_H_GEST_DIAR_PS_B2B, hay que coger la  fecha_de_baja_del_contrato_ps
                // a partir la fecha de la última actualizacion en t_ed_h_pdtweb_pm_b2b y la ultima fecha de ejecucion en T_ED_H_GEST_DIAR_PS_B2B
                //////strSql = "select A.comercializadora, A.cups20, A.cups22, A.contrato_ps, A.fecha_desde, A.fecha_hasta, A.mes, A.estado, A.distribuidora, "
                //////+ " A.multipunto, A.ritmo_facturacion, A.tco_segm_back, A.fecha_informe, A.fec_act, A.cod_carga, A.id_pte_web ,fecha_de_baja_del_contrato_ps as Final, fecha_de_alta_del_contrato_ps "
                //////+ " from "
                //////+ " ( "
                //////+ "     SELECT comercializadora, t_ed_h_pdtweb_pm_b2b.cups20, cups22, contrato_ps, fecha_desde, fecha_hasta, mes, estado, t_ed_h_pdtweb_pm_b2b.distribuidora, "
                //////+ "     t_ed_h_pdtweb_pm_b2b.multipunto, t_ed_h_pdtweb_pm_b2b.ritmo_facturacion, tco_segm_back, fecha_informe, t_ed_h_pdtweb_pm_b2b.fec_act, t_ed_h_pdtweb_pm_b2b.cod_carga "
                //////+ "     , id_pte_web "
                //////+ "     from ed_owner.t_ed_h_pdtweb_pm_b2b "
                //////+ "     where t_ed_h_pdtweb_pm_b2b.fecha_informe >= '" + fecha_max.AddDays(-10).ToString("yyyy-MM-dd") + "'"
                //////+ "     and t_ed_h_pdtweb_pm_b2b.fecha_informe < '" + fecha_max.ToString("yyyy-MM-dd") + "'"
                //////+ " ) as A "
                //////+ " left join ed_owner.T_ED_H_GEST_DIAR_PS_B2B "
                //////+ " on ed_owner.T_ED_H_GEST_DIAR_PS_B2B.cups20 = A.cups20 "
                //////+ " and fecha_registro_entidad_ps >= '" + fecha_max.AddDays(-10).ToString("yyyy-MM-dd") + "'"
                //////+ " and fecha_de_inicio_de_version_del_contrato_ps <= A.fecha_hasta "
                //////+ " and (fecha_de_fin_de_version_del_contrato_ps >= A.fecha_desde or fecha_de_fin_de_version_del_contrato_ps is null) "
                //////+ " group by  A.comercializadora, A.cups20, A.cups22, A.contrato_ps, A.fecha_desde, A.fecha_hasta, A.mes, A.estado, A.distribuidora,  "
                //////+ " A.multipunto, A.ritmo_facturacion, A.tco_segm_back, A.fecha_informe, A.fec_act, A.cod_carga, A.id_pte_web ,fecha_de_baja_del_contrato_ps, fecha_de_alta_del_contrato_ps "
                //////+ " order by A.cups20 asc, A.cups22 asc, A.fecha_desde asc, A.fecha_hasta asc, A.estado asc, A.distribuidora asc, A.fecha_informe desc ";

                strSql = "select cd_cups_20, id_contrato, fh_desde,fh_hasta,cd_estado_medida, de_estado_medida,de_marca_back,cod_carga,fh_act,fh_modificacion "
                + " from ed_owner.t_ed_h_kee_medidas_control "
                + " where UPPER(de_estado_medida) not like '%DESECHADA%'"
                + " and de_circuito_fact = 'BTN'"
                + " order by fh_hasta asc";

                ficheroLog.Add(strSql);
                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    PendienteMedida_B2B c = new PendienteMedida_B2B();
                    ////c.comercializadora = r["comercializadora"].ToString();
                    c.cups20 = r["cd_cups_20"].ToString();
                    ////c.cups22 = r["cups22"].ToString();
                    c.contrato_ps = r["id_contrato"].ToString();
                    if (r["fh_desde"] != System.DBNull.Value)
                    {
                        c.fecha_desde = Convert.ToDateTime(r["fh_desde"]);
                    }
                    if (r["fh_hasta"] != System.DBNull.Value)
                    {
                        c.fecha_hasta = Convert.ToDateTime(r["fh_hasta"]);
                    }
                    //////c.fecha_hasta = Convert.ToDateTime(r["fecha_hasta"]);
                    //////c.mes = Convert.ToInt32(r["mes"]);
                    ///
                    c.cod_estado = r["cd_estado_medida"].ToString();
                    c.estado = r["de_estado_medida"].ToString();
                    c.cod_carga = Convert.ToInt32(r["cod_carga"]);

                    ////c.fecha_informe = Convert.ToDateTime(r["fecha_informe"]);
                    c.fec_act = Convert.ToDateTime(r["fh_act"]);
                    if (r["fh_modificacion"] != System.DBNull.Value)
                    {
                        c.fecha_modificacion = Convert.ToDateTime(r["fh_modificacion"]);
                    }

                    //////if (r["Final"] != System.DBNull.Value)
                    //////{
                    //////    c.fecha_fin_KEE = Convert.ToDateTime(r["Final"]);
                    //////}

                    //////if (r["fecha_de_alta_del_contrato_ps"] != System.DBNull.Value)
                    //////{
                    //////    c.fec_alta_kee = Convert.ToDateTime(r["fecha_de_alta_del_contrato_ps"]);
                    //////}

                    ////////Paco 09/02/2023 Multipunto
                    //////if (r["multipunto"] != System.DBNull.Value)
                    //////{
                    //////    if (r["multipunto"].ToString() == "S")
                    //////    {
                    //////        c.multipunto = true;
                    //////    }
                    //////    else
                    //////    {
                    //////        c.multipunto = false;
                    //////    }
                    //////}
                    //////////////

                    List<PendienteMedida_B2B> t;
                    if (!d.TryGetValue(c.cups20, out t))
                    {
                        t = new List<PendienteMedida_B2B>();
                        t.Add(c);
                        d.Add(c.cups20, t);
                    }
                    else
                        t.Add(c);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception ex)
            {
                ficheroLog.addError("Carga: " + ex.Message);
                return null;
            }
        }

        private Dictionary<string, List<PendienteMedida_B2B>> CargaFechaBajaSAP()
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";

            ////fecha_ControlCargas = MaxFechaInformeGestDiarioB2B();
            Dictionary<string, List<PendienteMedida_B2B>> d = new Dictionary<string, List<PendienteMedida_B2B>>();

            try
            {

                strSql = "select A.id_crto_ext, max(A.fh_baja) as Fecha "
                + " from("
                + "     SELECT  id_crto_ext, max(cd_sec_crto), fh_baja"
                + "     from ed_owner.t_ed_h_sap_crto_front"
                + "     where de_marca_back = 'OPERACIONES B2B'"
                + "     group by id_crto_ext, fh_baja"
                + " ) as A"
                + " group by A.id_crto_ext";

 
                ficheroLog.Add(strSql);
                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    PendienteMedida_B2B c = new PendienteMedida_B2B();
                    ////c.comercializadora = r["comercializadora"].ToString();
                    c.id_crto_ext = r["id_crto_ext"].ToString();
                   
                    if (r["Fecha"] != System.DBNull.Value)
                    {
                        c.fecha_baja_sap = Convert.ToDateTime(r["Fecha"]);
                    }
                  

                    List<PendienteMedida_B2B> t;
                    if (!d.TryGetValue(c.id_crto_ext, out t))
                    {
                        t = new List<PendienteMedida_B2B>();
                        t.Add(c);
                        d.Add(c.id_crto_ext, t);
                    }
                    else
                        t.Add(c);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception ex)
            {
                ficheroLog.addError("Carga: " + ex.Message);
                return null;
            }
        }


        private Dictionary<string, List<PendienteMedida_B2B>>RelacionIncidenciasCUPS()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string ControlCUPS = "";


            ////fecha_ControlCargas = MaxFechaInformeGestDiarioB2B();
            Dictionary<string, List<PendienteMedida_B2B>> d = new Dictionary<string, List<PendienteMedida_B2B>>();

            try
            {

                //////strSql = "select A.id_crto_ext, max(A.fh_baja) as Fecha "
                //////+ " from("
                //////+ "     SELECT  id_crto_ext, max(cd_sec_crto), fh_baja"
                //////+ "     from ed_owner.t_ed_h_sap_crto_front"
                //////+ "     where de_marca_back = 'OPERACIONES B2B'"
                //////+ "     group by id_crto_ext, fh_baja"
                //////+ " ) as A"
                //////+ " group by A.id_crto_ext";


                strSql = " select cups, Mes_pendiente,area, incidencia, estado_incidencia, fecha_apertura, prioridad_negocio, titulo, e_s_estado, Reincidente " 
                    + " from Relacion_INC_CUPS"  
                    + " GROUP BY cups, Mes_pendiente,area, incidencia, estado_incidencia, fecha_apertura, prioridad_negocio, titulo, e_s_estado, Reincidente "
                    + " order by cups, Mes_pendiente,Fecha_apertura asc";


                ficheroLog.Add(strSql);

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    if (ControlCUPS != r["cups"].ToString()+ r["Mes_pendiente"].ToString())
                    {
                        ControlCUPS = r["cups"].ToString()+ r["Mes_pendiente"].ToString();

                        PendienteMedida_B2B c = new PendienteMedida_B2B();
                        ////c.comercializadora = r["comercializadora"].ToString();
                        c.cups = r["cups"].ToString();
                        c.mes_pendiente= r["Mes_pendiente"].ToString();
                        c.area = r["area"].ToString();
                        c.incidencia = r["incidencia"].ToString();
                        if (r["estado_incidencia"] != System.DBNull.Value)
                        {
                            c.estado_incidencia = r["estado_incidencia"].ToString();
                        }    
                        if (r["fecha_apertura"] != System.DBNull.Value)
                        {
                            c.fecha_apertura = Convert.ToDateTime(r["fecha_apertura"]);
                        }
                        if (r["prioridad_negocio"] != System.DBNull.Value)
                        {
                            c.prioridad_negocio = r["prioridad_negocio"].ToString();
                        }
                        if (r["titulo"] != System.DBNull.Value)
                        {
                            c.titulo = r["titulo"].ToString();
                        }
                        if (r["e_s_estado"] != System.DBNull.Value)
                        {
                            c.e_s_estado = r["e_s_estado"].ToString();
                        }
                        if (r["Reincidente"] != System.DBNull.Value)
                        {
                            c.Reincidente = r["Reincidente"].ToString();
                        }

                        List<PendienteMedida_B2B> t;
                        if (!d.TryGetValue(c.cups + c.mes_pendiente, out t))
                        {
                            t = new List<PendienteMedida_B2B>();
                            t.Add(c);
                            d.Add(c.cups + c.mes_pendiente, t);
                        }
                        else
                            t.Add(c);

                    }
                   
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception ex)
            {
                ficheroLog.addError("Carga: " + ex.Message);
                return null;
            }
        }

        private Dictionary<string, List<PendienteMedida_B2B>> Carga(DateTime fecha_max)
        {

            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";

            fecha_ControlCargas = MaxFechaInformeGestDiarioB2B();

            //////if (fecha_max.ToString("yyyy-MM-dd") != fecha_ControlCargas.ToString("yyyy-MM-dd"))
            //////{
            //////     System.Windows.Forms.MessageBox.Show("discrepancia ultima fecha de carga entre t_ed_h_pdtweb_pm_b2b y T_ED_H_GEST_DIAR_PS_B2B");
            //////}

            Dictionary<string, List<PendienteMedida_B2B>> d = new Dictionary<string, List<PendienteMedida_B2B>>();

            try
            {
                //////////strSql = "SELECT comercializadora, cups20, cups22, contrato_ps, fecha_desde, fecha_hasta,"
                //////////    + " mes, estado, distribuidora, multipunto, ritmo_facturacion, tco_segm_back, fecha_informe, fec_act,"
                //////////    + " cod_carga, id_pte_web"
                //////////    + " from ed_owner.t_ed_h_pdtweb_pm_b2b"
                //////////    + " where fec_act >= '" + fecha_max.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                ///DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd")

                //Paco- añadir fecha baja contrato KEE, hay que hacer left join con T_ED_H_GEST_DIAR_PS_B2B, hay que coger la  fecha_de_baja_del_contrato_ps
                // a partir la fecha de la última actualizacion en t_ed_h_pdtweb_pm_b2b y la ultima fecha de ejecucion en T_ED_H_GEST_DIAR_PS_B2B
                strSql = "select A.comercializadora, A.cups20, A.cups22, A.contrato_ps, A.fecha_desde, A.fecha_hasta, A.mes, A.estado, A.distribuidora, "
                + " A.multipunto, A.ritmo_facturacion, A.tco_segm_back, A.fecha_informe, A.fec_act, A.cod_carga, A.id_pte_web ,fecha_de_baja_del_contrato_ps as Final, fecha_de_alta_del_contrato_ps "
                + " from "
                + " ( "
                + "     SELECT comercializadora, t_ed_h_pdtweb_pm_b2b.cups20, cups22, contrato_ps, fecha_desde, fecha_hasta, mes, estado, t_ed_h_pdtweb_pm_b2b.distribuidora, "
                + "     t_ed_h_pdtweb_pm_b2b.multipunto, t_ed_h_pdtweb_pm_b2b.ritmo_facturacion, tco_segm_back, fecha_informe, t_ed_h_pdtweb_pm_b2b.fec_act, t_ed_h_pdtweb_pm_b2b.cod_carga "
                + "     , id_pte_web "
                + "     from ed_owner.t_ed_h_pdtweb_pm_b2b "
                ///+ "     where t_ed_h_pdtweb_pm_b2b.fec_act >= '" + fecha_max.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                //////+ "     where t_ed_h_pdtweb_pm_b2b.fecha_informe >= '" + DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd") + "'"
                ////// Hay problemas de sincronización, KEE sale todos los días, en cambio SAP no se ejecuta fines de semana y ¿Festivos?
                ////// A la hora de ir a buscar a KEE, buscamos la fecha maxima de SAP para sincronizar los dos
                //Paco 04/04/2024
                + "     where t_ed_h_pdtweb_pm_b2b.fecha_informe >= '" + fecha_max.AddDays(-10).ToString("yyyy-MM-dd") + "'"
                + "     and t_ed_h_pdtweb_pm_b2b.fecha_informe < '" + fecha_max.ToString("yyyy-MM-dd") + "'"
                + " ) as A "
                + " left join ed_owner.T_ED_H_GEST_DIAR_PS_B2B "
                + " on ed_owner.T_ED_H_GEST_DIAR_PS_B2B.cups20 = A.cups20 "
                ///+ " and fh_ejecucion >= '" + fecha_max.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                //////+ " and fecha_registro_entidad_ps >= '" + DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd") + "'"
                ////// Hay problemas de sincronización, KEE sale todos los días, en cambio SAP no se ejecuta fines de semana y ¿Festivos?
                ////// A la hora de ir a buscar a KEE, buscamos la fecha maxima de SAP para sincronizar los dos
                //Paco 04/04/2024
                + " and fecha_registro_entidad_ps >= '" + fecha_max.AddDays(-10).ToString("yyyy-MM-dd") + "'"
                //////+ " and  fecha_registro_entidad_ps < '" + fecha_max.ToString("yyyy-MM-dd") + "'"
                + " and fecha_de_inicio_de_version_del_contrato_ps <= A.fecha_hasta "
                + " and (fecha_de_fin_de_version_del_contrato_ps >= A.fecha_desde or fecha_de_fin_de_version_del_contrato_ps is null) "
                + " group by  A.comercializadora, A.cups20, A.cups22, A.contrato_ps, A.fecha_desde, A.fecha_hasta, A.mes, A.estado, A.distribuidora,  "
                + " A.multipunto, A.ritmo_facturacion, A.tco_segm_back, A.fecha_informe, A.fec_act, A.cod_carga, A.id_pte_web ,fecha_de_baja_del_contrato_ps, fecha_de_alta_del_contrato_ps "
                + " order by A.cups20 asc, A.cups22 asc, A.fecha_desde asc, A.fecha_hasta asc, A.estado asc, A.distribuidora asc, A.fecha_informe desc ";

                ficheroLog.Add(strSql);
                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    PendienteMedida_B2B c = new PendienteMedida_B2B();
                    c.comercializadora = r["comercializadora"].ToString();
                    c.cups20 = r["cups20"].ToString();
                    c.cups22 = r["cups22"].ToString();
                    c.contrato_ps = r["contrato_ps"].ToString();
                    if (r["fecha_desde"] != System.DBNull.Value)
                    { 
                        c.fecha_desde = Convert.ToDateTime(r["fecha_desde"]);
                    }
                    if (r["fecha_hasta"] != System.DBNull.Value)
                    { 
                        c.fecha_hasta = Convert.ToDateTime(r["fecha_hasta"]);
                     }
                    //////c.fecha_hasta = Convert.ToDateTime(r["fecha_hasta"]);
                    c.mes = Convert.ToInt32(r["mes"]);
                    c.estado = r["estado"].ToString();
                    c.cod_carga = Convert.ToInt32(r["cod_carga"]);
                    c.fecha_informe = Convert.ToDateTime(r["fecha_informe"]);
                    c.fec_act = Convert.ToDateTime(r["fec_act"]);

                    if (r["Final"] != System.DBNull.Value)
                    {
                        c.fecha_fin_KEE = Convert.ToDateTime(r["Final"]);
                    }

                    if (r["fecha_de_alta_del_contrato_ps"] != System.DBNull.Value)
                    {
                        c.fec_alta_kee = Convert.ToDateTime(r["fecha_de_alta_del_contrato_ps"]);
                    }
                    //////else {
                    //////    c.fecha_fin_KEE = DBNull.Value;
                    //////}

                    //Paco 09/02/2023 Multipunto
                    if (r["multipunto"] != System.DBNull.Value)
                    {
                        if (r["multipunto"].ToString() == "S")
                        {
                            c.multipunto = true;
                        }
                        else
                        {
                            c.multipunto = false;
                        }
                    }
                    ////////

                    List<PendienteMedida_B2B> t;
                    if (!d.TryGetValue(c.cups20, out t))
                    {
                        t = new List<PendienteMedida_B2B>();
                        t.Add(c);
                        d.Add(c.cups20, t);
                    }
                    else
                        t.Add(c);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception ex)
            {
                ficheroLog.addError("Carga: " + ex.Message);
                return null;
            }
        }

        private DateTime CalculaFechaDesdeinformeMaxPendiente()
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

            strSql = "select max(fec_act) as max_fecha from t_ed_h_sap_pendiente_facturar";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["max_fecha"] != System.DBNull.Value)
                    date = Convert.ToDateTime(r["max_fecha"]);
                break;
            }
            db.CloseConnection();
            return date;
        }

        private DateTime MaxFechaInforme()
        {
            // Buscamos la última fecha de publicación del informe

            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";
            DateTime fecha_max = new DateTime();

            try
            {
                strSql = "select max(fecha_informe) as max_fecha from ed_owner.t_ed_h_pdtweb_pm_b2b";


                ficheroLog.Add(strSql);
                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha_max = Convert.ToDateTime(r["max_fecha"]);
                }
                db.CloseConnection();

                return fecha_max;
            }
            catch (Exception ex)
            {
                ficheroLog.addError("MaxFechaInforme: " + ex.Message);
                return DateTime.Now;
            }

        }

        private DateTime MaxFechaInformeGestDiarioB2B()
        {
            // Buscamos la última fecha de publicación del informe

            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";
            DateTime fecha_max = new DateTime();

            try
            {
                ///strSql = "select max(fh_ejecucion) as max_fecha from ed_owner.T_ED_H_GEST_DIAR_PS_B2B";
                strSql = "select max(fec_act) as max_fecha from ed_owner.t_ed_h_sap_pendiente_facturar";

                ficheroLog.Add(strSql);
                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha_max = Convert.ToDateTime(r["max_fecha"]);
                }
                db.CloseConnection();

                return fecha_max;
            }
            catch (Exception ex)
            {
                ficheroLog.addError("MaxFechaInforme: " + ex.Message);
                return DateTime.Now;
            }
        }

        public List<string> GetCups(string cups20, DateTime fecha_desde, DateTime fecha_hasta,  DateTime fec_act)
        {
            List<string> lista = new List<string>();

            List<PendienteMedida_B2B> o;
            if (dic.TryGetValue(cups20, out o))
            {

                //////lista =
                //////    o.FindAll(z => z.fecha_desde == fecha_desde
                //////    && z.fecha_hasta == fecha_hasta && Convert.ToDateTime(fec_act.ToString("yyyy-MM-dd hh:mm")) > Convert.ToDateTime(z.fecha_informe.ToString("yyyy-MM-dd hh:mm"))  ).Select(z => z.estado).ToList();


                //////lista =
                //////   o.FindAll(z => z.fecha_desde == fecha_desde
                //////   && z.fecha_hasta == fecha_hasta ).Select(z => z.estado).ToList();


                lista =
                   o.FindAll(z => z.fecha_desde == fecha_desde
                   && z.fecha_hasta == fecha_hasta).Select(z => z.estado).ToList();

                //////lista =
                //////        o.FindAll(z => z.fecha_desde == fecha_desde
                //////        && z.fecha_hasta == fecha_hasta).Select(z => z.estado + "-" + Convert.ToString(z.fecha_informe)).ToList();
                ///
                //Convert.ToString(z.fecha_informe.ToString("yyyy-MM-dd"))
            }
            return lista;
        }

        public List<string> GetCupsHoja(string cups20, DateTime fecha_desde, DateTime fecha_hasta)
        {
            List<string> lista = new List<string>();

            List<PendienteMedida_B2B> o;
            if (dic.TryGetValue(cups20, out o))
            {
                lista =
                    o.FindAll(z => z.fecha_desde == fecha_desde
                    && z.fecha_hasta == fecha_hasta ).Select(z => z.estado).ToList();

                //////lista =
                //////        o.FindAll(z => z.fecha_desde == fecha_desde
                //////        && z.fecha_hasta == fecha_hasta).Select(z => z.estado + "-" + Convert.ToString(z.fecha_informe)).ToList();
            }
            return lista;
        }


        public List<string> GetCupsFinKEE(string cups20, DateTime fecha_desde, DateTime fecha_hasta)
        {
            List<string> lista = new List<string>();

            List<PendienteMedida_B2B> o;
            if (dic.TryGetValue(cups20, out o))
            {
                lista =  o.FindAll(z => z.fecha_desde == fecha_desde && z.fecha_hasta == fecha_hasta).Select(z => Convert.ToString(z.fecha_fin_KEE)).ToList();
            }
            return lista;
        }

        public List<string> GetCupsFechaBajaSAP(string id_crto_ext)
        {
            List<string> lista = new List<string>();

            List<PendienteMedida_B2B> o;
            if (dic.TryGetValue(id_crto_ext, out o))
            {
                lista = o.FindAll(z => z.id_crto_ext == id_crto_ext).Select(z => Convert.ToString(z.fecha_baja_sap)).ToList();
            }
            return lista;
        }

        public List<string> GetCupsFinKEENew(string cups20, DateTime fec_act)
        {
            List<string> lista = new List<string>();

            List<PendienteMedida_B2B> o;
            if (dic.TryGetValue(cups20, out o))
            {
                lista = o.FindAll(z => z.cups20 == cups20 && z.fecha_informe.ToString("yyyy-MM-dd") == fec_act.AddDays(-1).ToString("yyyy-MM-dd")).Select(z => Convert.ToString(z.fecha_fin_KEE)).ToList();
            }
            return lista;
        }

        public List<string> GetCupsFechaAltaKEE(string cups20, DateTime fec_act)
        {
            List<string> lista = new List<string>();

            List<PendienteMedida_B2B> o;
            if (dic.TryGetValue(cups20, out o))
            {
                lista = o.FindAll(z => z.cups20 == cups20 && z.fecha_informe.ToString("yyyy-MM-dd") == fec_act.AddDays(-1).ToString("yyyy-MM-dd")).Select(z => Convert.ToString(z.fec_alta_kee)).ToList();
            }
            return lista;
        }

        public List<string> GetCupsMultipunto(string cups20)
        {
            List<string> lista = new List<string>();

            List<PendienteMedida_B2B> o;
            if (dic.TryGetValue(cups20, out o))
            {
                lista = o.FindAll(z => z.cups20 == cups20 ).Select(z => Convert.ToString(z.multipunto)).ToList();
            }
            return lista;
        }

        public List<string> GetCupsDetalle(string cups20, DateTime fecha_desde, DateTime fecha_hasta, DateTime fh_act)
        {
            List<string> lista = new List<string>();
            int listado;
            string comparar;
            string[] cadena;

            List<PendienteMedida_B2B> o;
            if (dic.TryGetValue(cups20, out o))
            {

                //Coincide el trio recupero el estado de KEE --> Nuevas columnas de KEE en la hoja detalle
                lista =
                    o.FindAll(z => z.fecha_desde == fecha_desde
                    && z.fecha_hasta == fecha_hasta && z.fecha_informe.ToString("yyyy-MM-dd") == fh_act.AddDays(-1).ToString("yyyy-MM-dd")).Select(z => z.estado).ToList();

                if (lista.Count == 0) //No coincide el trio
                {
                    //Existen fechas parciales en KEE --> periodo fecha_desde/fecha_hasta comprendido parte en uno de KEE (z.fecha_desde/z.fecha_hasta)
                    //fecha desde KEE menor que fecha inicio y  fecha hasta KEE >= fecha desde !!SOLAPE POR LA IZQUIERDA!!
                    lista =
                    o.FindAll(z =>  z.fecha_desde < fecha_desde 
                    &&  z.fecha_hasta >= fecha_desde && z.fecha_hasta <= fecha_hasta && z.fecha_informe.ToString("yyyy-MM-dd") == fh_act.AddDays(-1).ToString("yyyy-MM-dd")).Select(z => z.estado + ";(" + Convert.ToString(z.fecha_desde.ToString("yyyyMMdd")) + "/" + Convert.ToString(z.fecha_hasta.ToString("yyyyMMdd")) + ")").ToList(); ;  //intercalado en z.fecha_desde

                    if (lista.Count == 0) //!!No hay SOLAPE POR LA IZQUIERDA!!
                    {
                        //fecha desde SAP menor que fecha desde de kronos y  fecha hasta SAP < fecha hasta de kronos  !!SOLAPE POR LA DERECHA!!
                        lista =
                        o.FindAll(z =>  z.fecha_desde < fecha_hasta 
                        &&  z.fecha_hasta > fecha_hasta && z.fecha_desde >= fecha_desde && z.fecha_informe.ToString("yyyy-MM-dd") == fh_act.AddDays(-1).ToString("yyyy-MM-dd")).Select(z => z.estado + ";(" + Convert.ToString(z.fecha_desde.ToString("yyyyMMdd")) + "/" + Convert.ToString(z.fecha_hasta.ToString("yyyyMMdd")) + ")").ToList(); //intercalado en z.fecha_hasta

                        if (lista.Count == 0) //!!No hay SOLAPE POR LA DERECHA!!
                        {      
                            lista =
                            o.FindAll(z => z.fecha_desde < fecha_desde
                            && fecha_hasta < z.fecha_hasta && z.fecha_informe.ToString("yyyy-MM-dd") == fh_act.AddDays(-1).ToString("yyyy-MM-dd")).Select(z => z.estado + ";(" + Convert.ToString(z.fecha_desde.ToString("yyyyMMdd")) + "/" + Convert.ToString(z.fecha_hasta.ToString("yyyyMMdd")) + ")").ToList(); //intercalado en z.fecha_hasta

                            if (lista.Count == 0) //No hay solapes de fecha, miro si están contenidas las fechas de KEE dentro de SAP
                            {
                                
                               lista =
                               o.FindAll(z => z.fecha_desde >= fecha_desde
                               && z.fecha_hasta <= fecha_hasta && z.fecha_informe.ToString("yyyy-MM-dd") == fh_act.AddDays(-1).ToString("yyyy-MM-dd")).Select(z => z.estado + ";(" + Convert.ToString(z.fecha_desde.ToString("yyyyMMdd")) + "/" + Convert.ToString(z.fecha_hasta.ToString("yyyyMMdd")) + ")").ToList(); // el periodo de KEE está contenido dentro del de SAP

                                if (lista.Count == 0)
                                {
                                    // En el caso de que en el pendiente de SAP haya un periodo pendiente en un estado de los que SI va a buscar a KEE y
                                    // no se encuentra en el pendiente de KEE, ni total ni parcialmente, se debería informar como el estado
                                    //"01.B04 Error Sistemas KEE-SAP - Recepción OL" que es el que identifica problemas en la generación del periodo 
                                    //ya sea por el envío de OL de SAP o por el proceso de KEE
                                    lista.Add("Discrepancia: No existen fechas en el informe pendiente de KEE para el periodo;(" + Convert.ToString(fecha_desde.ToString("yyyyMMdd")) + "/" + Convert.ToString(fecha_hasta.ToString("yyyyMMdd")) + ")");
                                }
                                else
                                {
                                    ///Periodo KEE contenido dentro del SAP, nos quedamos con el más antiguo (esta ordenado de forma asc)
                                    ///ejemplo: sap de 01/08/2023 a 31/12/2023 y tengo en kronos 1/11/2023 a 30/11/2023  - 01/12/2023- 31/12/2023
                                    /// z.fechadesde>= 01/08/2023 y z.fechahasta<= 31/12/2023  
                                    comparar = "";
                                    listado = lista.Count;
                                    cadena = lista[0].Split(';');
                                    comparar = cadena[1];

                                    for (int i = listado-1; i > 0; i--)
                                    {
                                        cadena = lista[i].Split(';');
                                        if (cadena[1] != comparar)
                                        {
                                            lista.RemoveAt(i);
                                        }
                                    }
                                    listado = lista.Count-1;
                                    for (int i = 0 ; i <= listado; i++)
                                    {
                                        lista[i] = "Discrepancia: Periodos del pendiente de KEE contenidos en las fechas de SAP;" + lista[i];
                                    }
                                    //lista[0] = "Discrepancia: Periodos del pendiente de KEE contenidos en las fechas de SAP;" + lista[0];
                                } //FIN if (lista.Count == 0)  KEE CONTENIDO DENTRO DE SAP
                            }
                            else
                            {
                                //Intercalado parcialmente
                                comparar = "";
                                listado = lista.Count;
                                cadena = lista[0].Split(';');
                                comparar = cadena[1];

                                for (int i = listado - 1; i > 0; i--)
                                {
                                    cadena = lista[i].Split(';');
                                    if (cadena[1] != comparar)
                                    {
                                        lista.RemoveAt(i);
                                    }
                                }
                                listado = lista.Count - 1;
                                for (int i = 0; i <= listado; i++)
                                {
                                    lista[i] = "Discrepancia: Periodo SAP contenido dentro de un informe pendiente de KEE;" + lista[i];
                                }
                                //////lista[0] = "Discrepancia: Periodo SAP contenido dentro de un informe pendiente de KEE;" + lista[0];
                            }//FIN if (lista.Count == 0) fechas de sap contenidas dentro de una fecha de kee
                        }
                        else
                        {
                            //Intercalado parcialmente
                            comparar = "";
                            listado = lista.Count;
                            cadena = lista[0].Split(';');
                            comparar = cadena[1];

                            for (int i = listado - 1; i > 0; i--)
                            {
                                cadena = lista[i].Split(';');
                                if (cadena[1] != comparar)
                                {
                                    lista.RemoveAt(i);
                                }
                            }
                            listado = lista.Count - 1;
                            for (int i = 0; i <= listado; i++)
                            {
                                lista[i] = "Discrepancia: Periodo SAP encontrado parcialmente en el informe pendiente de KEE;" + lista[i];
                            }
                            //Intercalado parcialmente
                            //lista[0] = "Discrepancia: Periodo SAP encontrado parcialmente en el informe pendiente de KEE;" + lista[0];
                        } //FIN if (lista.Count == 0) SOLAPE POR LA DERECHA
                    }
                    else
                    {
                        //Intercalado parcialmente
                        comparar = "";
                        listado = lista.Count;
                        cadena = lista[0].Split(';');
                        comparar = cadena[1];

                        for (int i = listado - 1; i > 0; i--)
                        {
                            cadena = lista[i].Split(';');
                            if (cadena[1] != comparar)
                            {
                                lista.RemoveAt(i);
                            }
                        }
                        listado = lista.Count - 1;
                        for (int i = 0; i <= listado; i++)
                        {
                            lista[i] = "Discrepancia: Periodo SAP encontrado parcialmente en el informe pendiente de KEE;" + lista[i];
                        }
                        ////lista[0] = "Discrepancia: Periodo SAP encontrado parcialmente en el informe pendiente de KEE;" + lista[0];
                    } //FIN if (lista.Count == 0) SOLAPE POR LA IZQUIERDA

                } //FIN if (lista.Count == 0) //No coincide el trio
            }
            else
            {
                //No encuentra el CUPS de SAP dentro del listado de CUPS de Kronos
                lista.Add("Discrepancia: No existe el cups en el informe del pendiente de KEE");
            }

            return lista;

        }

        public List<string> GetCupsDetalle_BTN(string cups20, DateTime fecha_desde, DateTime fecha_hasta, DateTime fh_act)
        {
            List<string> lista = new List<string>();
            int listado;
            string comparar;
            string[] cadena;

            List<PendienteMedida_B2B> o;
            if (dic.TryGetValue(cups20, out o))
            {

                //Adriana : Buscar el periodo solicitado por SAP en KEE: Buscaremos el primer periodo para el que fh_hasta KEE > FH_DESDE SAP

                // Nuevas columnas de KEE en la hoja detalle
                //////lista =
                //////    o.FindAll(z => z.fecha_hasta > fecha_desde && z.fecha_informe.ToString("yyyy-MM-dd") == fh_act.AddDays(-1).ToString("yyyy-MM-dd")).Select(z => z.estado + ";(" + Convert.ToString(z.fecha_desde.ToString("yyyyMMdd")) + "/" + Convert.ToString(z.fecha_hasta.ToString("yyyyMMdd")) + ")").ToList();


                //•	Si existen medidas posteriores no desechadas
                //////lista =
                //////    o.FindAll(z => z.fecha_hasta >= fecha_desde).Select(z => z.estado + ";(" + Convert.ToString(z.fecha_desde.ToString("yyyyMMdd")) + "/" + Convert.ToString(z.fecha_hasta.ToString("yyyyMMdd")) + ")" + ";" + Convert.ToString(z.fecha_modificacion.ToString("yyyyMMdd"))).ToList();


                lista =
                   o.FindAll(z => z.fecha_hasta > fecha_desde).Select(z => z.estado + ";(" + Convert.ToString(z.fecha_desde.ToString("yyyyMMdd")) + "/" + Convert.ToString(z.fecha_hasta.ToString("yyyyMMdd")) + ")" + ";" + Convert.ToString(z.fecha_modificacion.ToString("yyyyMMdd"))).ToList();


                if (lista.Count == 0) //No encuentra nada
                {

                    //No encuentra el CUPS de SAP dentro del listado de CUPS de Kronos
                    lista.Add("Discrepancia: No existen fechas en el informe pendiente de KEE para el periodo;(" + Convert.ToString(fecha_desde.ToString("yyyyMMdd")) + "/" + Convert.ToString(fecha_hasta.ToString("yyyyMMdd")) + ")");


                } //FIN if (lista.Count == 0) //No coincide el trio
                else
                {
                    ///Periodo KEE  nos quedamos con el más antiguo (esta ordenado de forma asc)
                    ///ejemplo: sap de 01/08/2023 a 31/12/2023 y tengo en kronos 1/11/2023 a 30/11/2023  - 01/12/2023- 31/12/2023
                    /// z.fechadesde>= 01/08/2023 y z.fechahasta<= 31/12/2023  
                    comparar = "";
                    listado = lista.Count;
                    cadena = lista[0].Split(';');
                    comparar = cadena[1];

                    for (int i = listado - 1; i > 0; i--)
                    {
                        cadena = lista[i].Split(';');
                        if (cadena[1] != comparar)
                        {
                            lista.RemoveAt(i);
                        }
                    }
                    //////listado = lista.Count - 1;
                    //////for (int i = 0; i <= listado; i++)
                    //////{
                    //////    lista[i] = "Discrepancia: Periodos del pendiente de KEE contenidos en las fechas de SAP;" + lista[i];
                    //////}

                }
            }
            else
            {
                //No encuentra el CUPS de SAP dentro del listado de CUPS de Kronos para BTN
                lista.Add("Discrepancia: No existe el cups en el informe del pendiente de KEE");
            }

            return lista;

        }

        public List<string> GetCupsDetalle_BTNInicioFinFacturada(string cups20)
        {
            List<string> lista = new List<string>();
            int listado;
            string comparar;
            string[] cadena;

            List<PendienteMedida_B2B> o;
            if (dic.TryGetValue(cups20, out o))
            {

                lista =
                    o.FindAll(z => z.cod_estado == "802").Select(z => z.estado + ";(" + Convert.ToString(z.fecha_desde.ToString("yyyyMMdd")) + "/" + Convert.ToString(z.fecha_hasta.ToString("yyyyMMdd")) + ")" ).ToList();


                if (lista.Count == 0) //No hay facturadas
                {

                    //No encuentra el CUPS de SAP dentro del listado de CUPS de Kronos
                    //lista.Add("Discrepancia: No existen fechas en el informe pendiente de KEE para el periodo;(" + Convert.ToString(fecha_desde.ToString("yyyyMMdd")) + "/" + Convert.ToString(fecha_hasta.ToString("yyyyMMdd")) + ")");


                } //FIN if (lista.Count == 0) //No coincide el trio
                else
                {
               
                    comparar = "";
                    listado = lista.Count;

                    //He obtenido todas las "Facturada", pero lo hago en orden FH_HASTA ascendente, por lo tanto me tengo que quedar con la última recuperada (la de fechas más altas)
                    // la ultima fecha facturada recuperada es la del indice lista.Count-1

                    listado = lista.Count;

                    //Empiezo por el final , si he recuperado 6 registros, el último es el de indice 5, borro de 4 para ab8ajo

                    for (int i = listado - 2; i >= 0; i--)
                    {
                           lista.RemoveAt(i);

                    }

                }
            }
            else
            {
                //No encuentra el CUPS de SAP dentro del listado de CUPS de Kronos
               /// lista.Add("Discrepancia: No existe el cups en el informe del pendiente de KEE");
            }

            return lista;

        }
    }
}
