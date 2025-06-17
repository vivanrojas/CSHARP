using EndesaBusiness.contratacion;
using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using EndesaEntity.cnmc.V21_2019_12_17;
using EndesaEntity.facturacion;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Graph;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion.redshift
{
       
    public class Pendiente : EndesaEntity.medida.Pendiente
    {
        utilidades.Param param;
        utilidades.Seguimiento_Procesos ss_pp;
        logs.Log ficheroLog;

        Dictionary<DateTime, List<EndesaEntity.medida.Pendiente>> dic_pendiente_hist_fecha;
        Dictionary<string, EndesaEntity.medida.Pendiente> dic_pendiente;
        Dictionary<string, DateTime> dic_dias_estado;
        Dictionary<string, DateTime> dic_dias_cups_subestado;

        Dictionary<string, EndesaEntity.facturacion.ResumenAgrupadaPendiente> dic_resumen_agrupadas_ES;
        Dictionary<string, EndesaEntity.facturacion.ResumenAgrupadaPendiente> dic_resumen_agrupadas_MTBTE;
        Dictionary<string, EndesaEntity.facturacion.ResumenAgrupadaPendiente> dic_resumen_agrupadas_BTN;
        

        public Pendiente()
        {
            param = new utilidades.Param("t_ed_h_sap_pendiente_param", servidores.MySQLDB.Esquemas.FAC);
            ss_pp = new utilidades.Seguimiento_Procesos();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Copia_Pendiente_BI");
            dic_pendiente = Carga();
        }

        private Dictionary<string, List<EndesaEntity.medida.Pendiente>> CargaPendienteTotal()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, List<EndesaEntity.medida.Pendiente>> d
                = new Dictionary<string, List<EndesaEntity.medida.Pendiente>>();

            try
            {
                strSql = " SELECT pend.empresa_titular AS EMPRESA,"
                    + " pend.cups13, "
                    + " pend.mes as aaaammPdte, pend.estado, pend.subestado"
                    + " FROM fact.t_ed_h_sap_pendiente_facturar pend"
                    + " ORDER BY pend.empresa_titular, "
                    + " pend.cups13, pend.mes ASC";

                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                    c.empresaTitular = r["EMPRESA"].ToString();
                    c.cups13 = r["cups13"].ToString().ToUpper();
                    c.aaaammPdte = Convert.ToInt32(r["aaaammPdte"]);
                    c.estado = r["estado"].ToString();
                    c.subsEstado = r["subestado"].ToString();

                    List<EndesaEntity.medida.Pendiente> o;
                    if (!d.TryGetValue(c.cups13, out o))
                    {
                        o = new List<EndesaEntity.medida.Pendiente>();
                        o.Add(c);
                        d.Add(c.cups13, o);
                    }
                    else
                        o.Add(c);



                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                ficheroLog.addError("CargaPendiente: " + e.Message);
                return null;
            }
        }

        public void CopiaDatos()
        {
            StringBuilder sb = new StringBuilder();
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";

            MySQLDB dbmy;
            MySqlCommand commandmy;

            bool firstOnly = true;
            int j = 0;
            int k = 0;
            int totalRegistros = 0;

            DateTime ultimaFechaCopiado = new DateTime();
            DateTime fechaInformeBI = new DateTime();

            utilidades.Fechas utilfecha = new Fechas();

            try
            {                

                ultimaFechaCopiado = UltimaActualizacionMySQL();
                //fechaInformeBI = UltimaActualizacionBI();

                if (ultimaFechaCopiado < DateTime.Now.Date)
                {
                    ss_pp.Update_Fecha_Inicio("Facturación", "Copia Pendiente BI", "Copia Pendiente BI");

                    //borrado_tabla();                    

                    ficheroLog.Add(Consulta(ultimaFechaCopiado));
                    db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                    command = new OdbcCommand(Consulta(ultimaFechaCopiado.AddMonths(-1)), db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        j++;
                        k++;

                        if (firstOnly)
                        {
                            sb = null;
                            sb = new StringBuilder();
                            sb.Append("REPLACE INTO t_ed_h_sap_pendiente_facturar");
                            sb.Append(" (cd_cups, id_instalacion, cl_stro, id_crto_ext, cl_crto_ext, cd_empr_distdora, fh_desde, fh_hasta,");
                            sb.Append("fh_periodo, cd_estado, cd_subestado, lg_multimedida, cd_empr_titular, cd_ritmo_fact,");
                            sb.Append("cd_segmento_ptg, fh_envio, fec_act, cod_carga, agora, tam) values ");

                            firstOnly = false;
                        }
                        #region campos

                        if (r["cd_cups"] != System.DBNull.Value)
                            sb.Append("('").Append(r["cd_cups"].ToString()).Append("',");
                        else
                            sb.Append("(null,");

                        if (r["id_instalacion"] != System.DBNull.Value)
                            sb.Append("'").Append(r["id_instalacion"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cl_stro"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cl_stro"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["id_crto_ext"] != System.DBNull.Value)
                            sb.Append("'").Append(r["id_crto_ext"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cl_crto_ext"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cl_crto_ext"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_empr_distdora"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_empr_distdora"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_desde"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_desde"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_hasta"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_hasta"]).ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_periodo"] != System.DBNull.Value)
                            sb.Append(r["fh_periodo"].ToString()).Append(",");
                        else
                            sb.Append("null,");

                        if (r["cd_estado"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_estado"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_subestado"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_subestado"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["lg_multimedida"] != System.DBNull.Value)
                            sb.Append("'").Append(r["lg_multimedida"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_empr_titular"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_empr_titular"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_ritmo_fact"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_ritmo_fact"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_segmento_ptg"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_segmento_ptg"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fh_envio"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fh_envio"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["fec_act"] != System.DBNull.Value)
                            sb.Append("'").Append(Convert.ToDateTime(r["fec_act"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cod_carga"] != System.DBNull.Value)
                            sb.Append(r["cod_carga"].ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        //AGORA                        
                        if (r["lg_agora"] != System.DBNull.Value)
                            sb.Append("'S',");
                        else
                            sb.Append("'N',");

                        if (r["nm_tam"] != System.DBNull.Value)
                            sb.Append(r["nm_tam"].ToString().Replace(",", ".")).Append("),");
                        else
                            sb.Append("null),");


                        #endregion



                        if (j == 100)
                        {
                            Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                            firstOnly = true;
                            dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                            commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                            commandmy.ExecuteNonQuery();
                            dbmy.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            j = 0;
                        }


                    }
                    db.CloseConnection();

                    if (j > 0)
                    {
                        Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                        firstOnly = true;
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                        commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        j = 0;
                    }

                    Construccion_Datos();

                    ss_pp.Update_Fecha_Fin("Facturación", "Copia Pendiente BI", "Copia Pendiente BI");

                }
                
                

            }
            catch(Exception ex)
            {
                ficheroLog.addError(ex.Message);
            }



        }

        public void CopiaDatosAlarmas()
        {
            StringBuilder sb = new StringBuilder();
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";

            MySQLDB dbmy;
            MySqlCommand commandmy;

            bool firstOnly = true;
            int j = 0;
            int k = 0;
            int totalRegistros = 0;

            DateTime ultimaFechaCopiadoPendienteBI = new DateTime();
            //DateTime fechaInformeBI = new DateTime();

            //utilidades.Fechas utilfecha = new Fechas();

            try
            {

                ultimaFechaCopiadoPendienteBI = UltimaActualizacionBI();
                //fechaInformeBI = UltimaActualizacionBI();

                if (ultimaFechaCopiadoPendienteBI > DateTime.ParseExact(param.GetValue("fecha_actualizacion_alarmas_pendiente_facturar"),"dd/MM/yyyy",CultureInfo.InvariantCulture))
                {
                    ss_pp.Update_Fecha_Inicio("Facturación", "Copia Alarmas Pendiente BI", "Copia Alarmas Pendiente BI");

                    //Borramos tabla de alarmas facturas pendientes (fact.t_ed_d_alarmasconfact_pendiente_facturar)
                    borrado_tabla_alarmas();

                    //Alarmas de contratos NO en Vigor

                    ficheroLog.Add(ConsultaAlarmasContratosNoEnVigor(ultimaFechaCopiadoPendienteBI));
                    db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                    command = new OdbcCommand(ConsultaAlarmasContratosNoEnVigor(ultimaFechaCopiadoPendienteBI), db.con);
                    command.CommandTimeout = 2000;
                    r = command.ExecuteReader();
                    
                    while (r.Read())
                    {
                        j++;
                        k++;

                        if (firstOnly)
                        {
                            sb = null;
                            sb = new StringBuilder();
                            sb.Append("REPLACE INTO t_ed_d_alarmasconfact_pendiente_facturar");
                            sb.Append(" (cemptitu, ccontrps, cd_cups_ext, cd_tp_alarma, dcomenta, finialar, ffinalar) values ");

                            firstOnly = false;
                        }
                        #region campos

                        if (r["cemptitu"] != System.DBNull.Value)
                            sb.Append("('").Append(r["cemptitu"].ToString()).Append("',");
                        else
                            sb.Append("(null,");

                        if (r["ccontrps"] != System.DBNull.Value)
                            sb.Append("'").Append(r["ccontrps"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_cups_ext"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_cups_ext"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_alarma"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_alarma"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["dcomenta"] != System.DBNull.Value)
                            sb.Append("'").Append(r["dcomenta"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["finialar"] != System.DBNull.Value)
                            sb.Append("'").Append(r["finialar"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["finialar"] != System.DBNull.Value)
                            sb.Append("'").Append(r["finialar"].ToString()).Append("'),");
                        else
                            sb.Append("null),");


                        #endregion



                        if (j == 100)
                        {
                            Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                            firstOnly = true;
                            dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                            commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                            commandmy.ExecuteNonQuery();
                            dbmy.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            j = 0;
                        }


                    }
                    db.CloseConnection();

                    if (j > 0)
                    {
                        Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                        firstOnly = true;
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                        commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        j = 0;
                    }

                    //Alarmas de contratos en Vigor
                    ficheroLog.Add(ConsultaAlarmasContratosEnVigor(ultimaFechaCopiadoPendienteBI));
                    db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                    command = new OdbcCommand(ConsultaAlarmasContratosEnVigor(ultimaFechaCopiadoPendienteBI), db.con);
                    command.CommandTimeout = 2000;
                    r = command.ExecuteReader();

                    while (r.Read())
                    {
                        j++;
                        k++;

                        if (firstOnly)
                        {
                            sb = null;
                            sb = new StringBuilder();
                            sb.Append("REPLACE INTO t_ed_d_alarmasconfact_pendiente_facturar");
                            sb.Append(" (cemptitu, ccontrps, cd_cups_ext, cd_tp_alarma, dcomenta, finialar, ffinalar) values ");

                            firstOnly = false;
                        }
                        #region campos

                        if (r["cemptitu"] != System.DBNull.Value)
                            sb.Append("('").Append(r["cemptitu"].ToString()).Append("',");
                        else
                            sb.Append("(null,");

                        if (r["ccontrps"] != System.DBNull.Value)
                            sb.Append("'").Append(r["ccontrps"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_cups_ext"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_cups_ext"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_tp_alarma"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_tp_alarma"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["dcomenta"] != System.DBNull.Value)
                            sb.Append("'").Append(r["dcomenta"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["finialar"] != System.DBNull.Value)
                            sb.Append("'").Append(r["finialar"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["finialar"] != System.DBNull.Value)
                            sb.Append("'").Append(r["finialar"].ToString()).Append("'),");
                        else
                            sb.Append("null),");


                        #endregion



                        if (j == 100)
                        {
                            Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                            firstOnly = true;
                            dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                            commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                            commandmy.ExecuteNonQuery();
                            dbmy.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            j = 0;
                        }


                    }
                    db.CloseConnection();

                    if (j > 0)
                    {
                        Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                        firstOnly = true;
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                        commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        j = 0;
                    }

                    ss_pp.Update_Fecha_Fin("Facturación", "Copia Alarmas Pendiente BI", "Copia Alarmas Pendiente BI");
                    //param.UpdateParameter("fecha_actualizacion_alarmas_pendiente_facturar", DateTime.Now.Date.ToString("dd/MM/yyyy"));
                    param.UpdateParameter("fecha_actualizacion_alarmas_pendiente_facturar", ultimaFechaCopiadoPendienteBI.ToString("dd/MM/yyyy"));
                }



            }
            catch (Exception ex)
            {
                ficheroLog.addError(ex.Message);
            }



        }

        private void borrado_tabla()
        {
            MySQLDB db;
            MySqlCommand command;
            
            string strSql = "";

            strSql = "delete from t_ed_h_sap_pendiente_facturar";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }
        private void borrado_tabla_alarmas()
        {
            MySQLDB db;
            MySqlCommand command;

            string strSql = "";

            strSql = "delete from t_ed_d_alarmasconfact_pendiente_facturar";
            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
             catch(Exception ex)
            {
                ficheroLog.addError(ex.Message);
            }
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

        private DateTime UltimaActualizacionBI()
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";
            DateTime fecha = new DateTime(2022, 01, 01);

            strSql = "SELECT max(fh_envio) AS fh_envio FROM ed_owner.t_ed_h_sap_pendiente_facturar";
            db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
            command = new OdbcCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())                
                if (r["fh_envio"] != System.DBNull.Value)
                    fecha = Convert.ToDateTime(r["fh_envio"]);
            db.CloseConnection();
            Console.WriteLine("Última fecha de copiado pendiente facturar BI: " + fecha.ToString("dd/MM/yyyy"));
            ficheroLog.Add("Última fecha de copiado pendeinte facturar BI: " + fecha.ToString("dd/MM/yyyy"));

            //21-10-2024 GUS: modificamos para que devuelva la fecha exacta de la actualización de la tabla en BI, previamente verificamos que esta función no tiene referencias
            //return fecha.AddDays(-4);
            return fecha;
        }

        private string Consulta(DateTime f)
        {
            string strSql = "";

            strSql = "SELECT p.cd_cups, i.id_instalacion, p.cl_stro, p.id_crto_ext, p.cl_crto_ext, p.cd_empr_distdora,"
                + " p.fh_desde, p.fh_hasta, p.fh_periodo, p.cd_estado, p.cd_subestado, p.lg_multimedida,"
                + " p.cd_empr_titular, p.cd_ritmo_fact, p.cd_segmento_ptg, p.fh_envio, p.fec_act, p.cod_carga,"
                + " i.nm_tam, i.lg_agora"
                + " FROM ed_owner.t_ed_h_sap_pendiente_facturar p"
                + " left outer join ed_owner.t_ed_h_sap_instalacion i on"
                + " i.cd_cups = p.cd_cups"
                + " where"
                + " fh_envio > '" + f.ToString("yyyy-MM-dd") + "'";

            return strSql;

        }
        private string ConsultaAlarmasContratosEnVigor(DateTime f)
        {
            string strSql = "";

            strSql = "select a.cemptitu, a.ccontrps, left(u.cd_cups_ext,20) as cd_cups_ext, a.cd_tp_alarma, a.dcomenta, a.finialar, a.ffinalar"
                + " from ed_owner.t_ed_d_alarmasconfact a"
                + " inner join ed_owner.t_ed_f_uvcrtos u on u.id_crto_ext = a.ccontrps and u.de_estado_crto IN ('EN VIGOR')"
                + " inner join ed_owner.t_ed_h_sap_instalacion i on i.cd_cups = u.cd_cups_ext"
                + " inner join ed_owner.t_ed_h_sap_pendiente_facturar p on p.cd_cups = i.cd_cups"
                + " where p.fh_periodo <= a.ffinalar and p.fh_envio = '" + f.ToString("yyyy-MM-dd") + "'"
                + " group by a.cemptitu, a.ccontrps, u.cd_cups_ext, a.cd_tp_alarma, a.dcomenta, a.finialar, a.ffinalar";

            return strSql;

        }
        private string ConsultaAlarmasContratosNoEnVigor(DateTime f)
        {
            string strSql = "";

            strSql = "select a.cemptitu, a.ccontrps, left(u.cd_cups_ext,20) as cd_cups_ext, a.cd_tp_alarma, a.dcomenta, a.finialar, a.ffinalar"
                + " from ed_owner.t_ed_d_alarmasconfact a"
                + " inner join ed_owner.t_ed_f_uvcrtos u on u.id_crto_ext = a.ccontrps and u.de_estado_crto NOT IN ('EN VIGOR')"
                + " inner join ed_owner.t_ed_h_sap_instalacion i on i.cd_cups = u.cd_cups_ext"
                + " inner join ed_owner.t_ed_h_sap_pendiente_facturar p on p.cd_cups = i.cd_cups"
                + " where p.fh_periodo <= a.ffinalar and p.fh_envio = '" + f.ToString("yyyy-MM-dd") + "'"
                + " group by a.cemptitu, a.ccontrps, u.cd_cups_ext, a.cd_tp_alarma, a.dcomenta, a.finialar, a.ffinalar";

            return strSql;

        }

        public void Construccion_Datos()
        {
            MySQLDB db;
            MySqlCommand command;            
            string strSql = "";

            try
            {
                strSql = "REPLACE INTO  t_ed_h_sap_pendiente_facturar_agrupado"
                    + " SELECT substr(cd_cups,1,20) as cd_cups, id_instalacion, cl_stro, id_crto_ext, cl_crto_ext,"
                    + " cd_empr_distdora, fh_desde, fh_hasta, fh_periodo, cd_estado, cd_subestado,"
                    + " lg_multimedida, cd_empr_titular, cd_ritmo_fact, cd_segmento_ptg, fh_envio,"
                    + " fec_act, cod_carga, TAM, agora, now()"
                    + " FROM t_ed_h_sap_pendiente_facturar p WHERE p.cd_subestado <> '03.I' and p.cd_subestado <> '03.C'"  //Añadimos este filtro a cd_subestado a petición de Ignacio Villar - 25/04/2024. Añadimos subestado 03.C a petición de Ignacio Villar - 19/09/2024
                    + " ORDER BY cd_cups, fh_periodo DESC, fh_hasta DESC";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "replace into t_ed_h_sap_tam_agora"
                + " SELECT g.cd_cups, g.id_instalacion, g.id_crto_ext,"
                + " g.cd_empr_titular, g.cd_segmento_ptg, g.tam, g.agora, g.fec_act"
                + " FROM t_ed_h_sap_pendiente_facturar_agrupado g"
                + " ORDER BY g.fec_act";
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch(Exception ex)
            {
                ficheroLog.addError(ex.Message);
                Console.WriteLine("Error en Construccion_Datos(): " + ex.Message);
            }

        }

        public void GeneraInformePendSAP(bool automatico)
        {
            FileInfo file;
            string ruta_salida_archivo = "";
            string ruta_salida_archivo_estados_descartados = "";

            //Eliminamos ficheros antiguos informe pendiente facturar SAP
            string[] listaArchivos = System.IO.Directory.GetFiles(automatico ? param.GetValue("ruta_salida_informe") : @"c:\Temp\",
                    param.GetValue("prefijo_informe") + "*.xlsx");

            for (int i = 0; i < listaArchivos.Length; i++)
            {
                file = new FileInfo(listaArchivos[i]);
                file.Delete();
            }

            //Eliminamos ficheros antiguos informe estados descartados pendiente facturar SAP
            listaArchivos = System.IO.Directory.GetFiles(automatico ? param.GetValue("ruta_salida_informe") : @"c:\Temp\",
                    param.GetValue("prefijo_informe_estados_descartados") + "*.xlsx");

            for (int i = 0; i < listaArchivos.Length; i++)
            {
                file = new FileInfo(listaArchivos[i]);
                file.Delete();
            }

            // Ruta salida informe pendiente SAP
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

            // Ruta salida informe estados descartados pendiente SAP
            if (automatico)
                ruta_salida_archivo_estados_descartados = param.GetValue("ruta_salida_informe")
                    + param.GetValue("prefijo_informe_estados_descartados")
                    + DateTime.Now.ToString("yyyyMMdd")
                    + param.GetValue("sufijo_informe_estados_descartados");
            else
                ruta_salida_archivo_estados_descartados = @"c:\Temp\"
                   + param.GetValue("prefijo_informe_estados_descartados")
                   + DateTime.Now.ToString("yyyyMMdd")
                   + param.GetValue("sufijo_informe_estados_descartados");

            //19/11/2024 GUS: añadimos nuevo parámetro, ruta salida fichero informe estados descartados 
            InformePendiente_BI_Facturacion(ruta_salida_archivo, ruta_salida_archivo_estados_descartados, automatico);
        }

        private void InformePendiente_BI_Facturacion(string ruta_salida_archivo, string ruta_salida_archivo_estados_descartados, bool automatico)
        {


            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            MySQLDB db_histo;
            MySqlCommand command_histo;
            MySqlDataReader r_histo;
            string strSql_histo = "";


            int c = 1;
            int f = 1;

            DateTime fecha_actual = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime fecha_registro = new DateTime();
            int meses_pdtes = 0;
            int aniomes = 0;

            bool tiene_complemento_a01 = false;
            bool sacar_portugal = true;
            bool tiene_num_dia_agrupacion = false;
            bool firstOnly = true;

            DateTime fd = new DateTime();
            DateTime fd_tam = new DateTime();
            DateTime udh = new DateTime();

            //MIO
            bool Pinto = true;
            string RangoInterno;
            string RangoPintoGris;
            int tamaño;
            string tamañoanterior = "";


            utilidades.Fechas utilfecha = new Fechas();

            try
            {                

                if (!automatico || (UltimaActualizacionMySQL().Date > 
                    ss_pp.GetFecha_FinProceso("Facturación", "Informe Pendiente BI", "Informe Pendiente BI").Date))
                {

                    if(automatico)
                        ss_pp.Update_Fecha_Inicio("Facturación", "Informe Pendiente BI", "Informe Pendiente BI");

                    /*
                    FileInfo plantillaExcel =  new FileInfo(System.Environment.CurrentDirectory +  param.GetValue("plantilla_informe_pendiente"));


                    FileInfo plantillaExcel =
                        new FileInfo(System.Environment.CurrentDirectory +
                        param.GetValue("plantilla_informe_pendiente"));

                    FileInfo fileInfo = new FileInfo(ruta_salida_archivo);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                    ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);

                    var workSheet = excelPackage.Workbook.Worksheets["Resumen ES"];
                    */

                    // *****************************************

                    
                    // Crea un fichero excel dinamicamente para informe pendiente SAP
                    //ruta_salida_archivo = "c:\\temp\\DetallePendFact_SAP_20231128_Inventarios_ES_PT_TAM.xlsx";
                    FileInfo fileInfo = new FileInfo(ruta_salida_archivo);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                    ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                    //Creo la primera hoja
                    var workSheet = excelPackage.Workbook.Worksheets.Add("Resumen ES");
                    var headerCells = workSheet.Cells[1, 1, 1, 17];
                    var headerFont = headerCells.Style.Font;

                    List<string> lista_empresas_ES = new List<string>();
                    lista_empresas_ES.Add("ES21");
                    lista_empresas_ES.Add("ES22");

                    List<string> lista_empresas_PT = new List<string>();
                    lista_empresas_PT.Add("PT1Q");

                    List<string> lista_segmentos_MT_BTE = new List<string>();
                    lista_segmentos_MT_BTE.Add("MT");
                    lista_segmentos_MT_BTE.Add("BTE");
                    lista_segmentos_MT_BTE.Add("AT");
                    lista_segmentos_MT_BTE.Add("MAT");

                    List<string> lista_segmentos_BTN = new List<string>();
                    lista_segmentos_BTN.Add("BTN");
                    lista_segmentos_BTN.Add("NULL");

                    // Tomamos lo últimos 5 días hábiles
                    // Si se lanza el listado en día fin de semana
                    // hay que quitar un día más porque el viernes todavía no se ha procesado

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

                    

                    fd_tam = utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil());
                    
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
                    dic_dias_cups_subestado = CargaDiasCUPSSubEstado();

                    dic_resumen_agrupadas_ES = new Dictionary<string, EndesaEntity.facturacion.ResumenAgrupadaPendiente>();
                    dic_resumen_agrupadas_MTBTE = new Dictionary<string, EndesaEntity.facturacion.ResumenAgrupadaPendiente>();
                    dic_resumen_agrupadas_BTN = new Dictionary<string, EndesaEntity.facturacion.ResumenAgrupadaPendiente>();

                    int dia = 0;
                    int dia_tam = 0;
                    int fila;
                    int columna;

                    ///dic_pendiente_hist_fecha = CargaPendienteHist_DesdeFecha(CalculaFechaDesdeinforme(), lista_empresas_ES);

                    fd_tam = utilfecha.UltimoDiaHabilAnterior(utilfecha.UltimoDiaHabil());

                    int InicioRango;
                    int pintoetiqueta;
                    bool blnPintoEtiqueta;

                    blnPintoEtiqueta = true;
                    int Hoja;
                    List<string> listaEmpresas = new List<string>();
                    List<string> listaSegmentos = new List<string>();

                    var allCells = workSheet.Cells[1, 1, 50, 50];
                    Color colorDeCeldaTitulo = ColorTranslator.FromHtml("#b3e3af");
                    Color colorDeCeldaCabecera = ColorTranslator.FromHtml("#c1dce6");
                    Color colorDeCeldaTotales = ColorTranslator.FromHtml("#e0f5c6");
                    Color colorDeCeldaGris = ColorTranslator.FromHtml("#e8eced");
                    Color colorAviso = ColorTranslator.FromHtml("#ff9a9b");

                    #region Resúmenes (ES, POR MT-BTE, POR BTN)
                    for (Hoja = 1; Hoja < 4; Hoja++)
                    {
                       
                        if (Hoja == 1) { //RESUMEN_ES

                            listaEmpresas = lista_empresas_ES;
                            listaSegmentos = null;
                            //dic_pendiente_hist_fecha = CargaPendienteHist_DesdeFecha(CalculaFechaDesdeinforme(), lista_empresas_ES);
                            dic_pendiente_hist_fecha = CargaPendienteHist_DesdeFecha(CalculaFechaDesdeinforme(), listaEmpresas);
                        }
                        if (Hoja == 2)
                        {
                            blnPintoEtiqueta = true;
                            Pinto = true;
                            dic_Totales_cups.Clear();
                            dic_Totales_tam.Clear();
                            listaEmpresas = lista_empresas_PT;
                            listaSegmentos = lista_segmentos_MT_BTE;
                            dic_pendiente_hist_fecha = CargaPendienteHist_PT_DesdeFecha(CalculaFechaDesdeinforme(), listaEmpresas, listaSegmentos);
                            dia = 0;
                            dia_tam = 0;
                            //CREO LA HOJA
                            workSheet = excelPackage.Workbook.Worksheets.Add("Resumen POR MT-BTE");
                            headerCells = workSheet.Cells[1, 1, 1, 17];
                            headerFont = headerCells.Style.Font;
                        }


                        if (Hoja == 3) {

                            blnPintoEtiqueta = true;
                            Pinto = true;
                            dic_Totales_cups.Clear();
                            dic_Totales_tam.Clear();
                            listaEmpresas = lista_empresas_PT;
                            listaSegmentos = lista_segmentos_BTN;

                            dic_pendiente_hist_fecha = CargaPendienteHist_PT_DesdeFecha(CalculaFechaDesdeinforme(), lista_empresas_PT, lista_segmentos_BTN);

                            dia = 0;
                            dia_tam = 0;

                            //CREO LA HOJA
                            workSheet = excelPackage.Workbook.Worksheets.Add("Resumen POR BTN");
                            headerCells = workSheet.Cells[1, 1, 1, 17];
                            headerFont = headerCells.Style.Font;
                        }

                        c = 9;
                        // PARTE NO ÁGORA ES ***************************************************
                        //Veo que etiquetas tengo en los dias
                        string[] Guardar;
                        Guardar = new string[150000];
                        int m;
                        m = 0;
                        foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                        {
                            dia++;
                            if (dia < 6)
                            {
                                int v = p.Value.Count;
                                for (int i = 0; i < v; i++)
                                {
                                    Guardar[m] = p.Value[i].cod_estado + "_" + p.Value[i].cod_subestado + "_" + p.Value[i].agora + "_" + p.Value[i].subsEstado;
                                    m++;
                                }
                            }
                        }
                       
                        string[] B = Guardar.Distinct().ToArray(); //Aquí ya tengo todas las etiquetas de los días

                        Array.Sort(B); //Ordeno Alfabeticamente

                    
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
                        workSheet.Cells["A3:A4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells["A3:A4"].Style.Border.BorderAround( ExcelBorderStyle.Thin);

                        workSheet.Cells["A3:A4"].Style.Fill.BackgroundColor.SetColor(colorDeCeldaCabecera);
                        workSheet.Cells["A3:A4"].Merge = true;
                        workSheet.Cells["A3:A4"].Style.WrapText = true;
                        //workSheet.Cells["A3:A4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        workSheet.Cells["A3:A4"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        //OfficeOpenXml.Style.ExcelHorizontalAlignement

                        workSheet.Cells[3, 2].Value = "RESPONSABLE";
                        workSheet.Cells[3, 2].Style.Font.Bold = true;
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

                        foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                        {
                            dia++;
                            if (dia < 6)
                            {
                                Console.WriteLine("Totales ES noAgora dia: " + p.Key.ToString("dd/MM/yyyy"));

                                f = 4;
                                c--;

                                workSheet.Cells[f, c].Value = p.Key;
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                workSheet.Cells[f, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                //colorDeCelda = ColorTranslator.FromHtml("#e6faf5");
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(colorDeCeldaCabecera);
                                workSheet.Cells[f,c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;

                                //for (int i = 0; i < B.Length - 1; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                for (int i = 1; i < B.Length ; i++) //Quito el 0  ya que en el primer item están los nulos que no he rellenado al poner limite 5000
                                    {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "01" & cadena[2] == "False") //cadena[2] == "False" ES NO AGORA
                                    {
                                        workSheet.Cells[f, 3].Value = cadena[1] + " " + cadena[3];
                                        workSheet.Cells[f, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f,c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                if (Pinto == true)
                                {
                                    tamaño = f;
                                    tamaño = tamaño - 1;
    
                                    if (tamaño >= 5 ) //empieza en fila 5
                                    {
                                        RangoInterno = "B5:B" + tamaño.ToString();
                                        workSheet.Cells[RangoInterno].Value = "Pendiente medida";
                                        workSheet.Cells[RangoInterno].Style.Font.Bold = true;
                                        workSheet.Cells[RangoInterno].Merge = true;
                                        workSheet.Cells[RangoInterno].Style.WrapText = true;
                                        workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                    }

                                }

                                workSheet.Cells[f, 3].Value = "TotalPendiente medida";
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                RangoPintoGris = "B" + f.ToString() + ":C" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                RangoPintoGris = "B" + f.ToString() + ":M" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaGris);

                                /*
                                if (Pinto == true)
                                {
                                    colorDeCelda = ColorTranslator.FromHtml("#f0f4f5"); //Color GRIS
                                    RangoInterno = "B" + f.ToString() + ":M" + f.ToString();
                                    workSheet.Cells[RangoInterno].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    workSheet.Cells[RangoInterno].Style.Fill.BackgroundColor.SetColor(colorDeCelda);
                                }
                                */
                                totales_noagora_01 = Total_Pendiente(noAgora, p.Key, listaEmpresas, listaSegmentos, "01", null);
                                workSheet.Cells[f, c].Value = totales_noagora_01;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                f++;
                                tamañoanterior = f.ToString();

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "02" & cadena[2] == "False")
                                    {
                                        workSheet.Cells[f, 3].Value = cadena[1] + " " + cadena[3];
                                        workSheet.Cells[f, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                if (Pinto == true)
                                {
                                    RangoInterno = "B" + tamañoanterior;
                                    tamaño = f;
                                    tamaño = tamaño - 1;
                                    
                                    if (tamaño >= Int32.Parse(tamañoanterior)) 
                                    {
                                        RangoInterno = RangoInterno + ":B" + tamaño.ToString();
                                        workSheet.Cells[RangoInterno].Value = "Orden de Cálculo Calculable";
                                        workSheet.Cells[RangoInterno].Style.Font.Bold = true;
                                        workSheet.Cells[RangoInterno].Merge = true;
                                        workSheet.Cells[RangoInterno].Style.WrapText = true;
                                        workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    }
                                }

                                workSheet.Cells[f, 3].Value = "Total Orden de Cálculo Calculable";
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                RangoPintoGris = "B" + f.ToString() + ":C" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                RangoPintoGris = "B" + f.ToString() + ":M" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaGris);

                                totales_noagora_02 = Total_Pendiente(noAgora, p.Key, listaEmpresas, listaSegmentos, "02", null);
                                workSheet.Cells[f, c].Value = totales_noagora_02;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;
                                tamañoanterior = f.ToString();

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "03" & cadena[2] == "False")
                                    {
                                        workSheet.Cells[f, 3].Value = cadena[1] + " " + cadena[3];
                                        workSheet.Cells[f, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                if (Pinto == true)
                                {
                                    RangoInterno = "B" + tamañoanterior;
                                    tamaño = f;
                                    tamaño = tamaño - 1;
                                    
                                    if (tamaño >= Int32.Parse(tamañoanterior)) 
                                    {
                                        RangoInterno = RangoInterno + ":B" + tamaño.ToString();
                                        workSheet.Cells[RangoInterno].Value = "DC Generado sin DI";
                                        workSheet.Cells[RangoInterno].Style.Font.Bold = true;
                                        workSheet.Cells[RangoInterno].Merge = true;
                                        workSheet.Cells[RangoInterno].Style.WrapText = true;
                                        workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    }
                                }

                                workSheet.Cells[f, 3].Value = "Total DC Generado sin DI";
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                RangoPintoGris = "B" + f.ToString() + ":C" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                RangoPintoGris = "B" + f.ToString() + ":M" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaGris);
                                totales_noagora_03 = Total_Pendiente(noAgora, p.Key, listaEmpresas, listaSegmentos, "03", null);
                                workSheet.Cells[f, c].Value = totales_noagora_03;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;
                                tamañoanterior = f.ToString();

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "04" & cadena[2] == "False")
                                    {
                                        workSheet.Cells[f, 3].Value = cadena[1] + " " + cadena[3];
                                        workSheet.Cells[f, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                if (Pinto == true)
                                {
                                    RangoInterno = "B" + tamañoanterior;
                                    tamaño = f;
                                    tamaño = tamaño - 1;
                                   
                                    if (tamaño >= Int32.Parse(tamañoanterior))
                                    {
                                        RangoInterno = RangoInterno + ":B" + tamaño.ToString();
                                        workSheet.Cells[RangoInterno].Value = "Doc. Impresión Apartado";
                                        workSheet.Cells[RangoInterno].Style.Font.Bold = true;
                                        workSheet.Cells[RangoInterno].Merge = true;
                                        workSheet.Cells[RangoInterno].Style.WrapText = true;
                                        workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    }
                                }

                                workSheet.Cells[f, 3].Value = "Total Doc. Impresión Apartado";
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                RangoPintoGris = "B" + f.ToString() + ":C" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                RangoPintoGris = "B" + f.ToString() + ":M" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaGris);
                                totales_noagora_04 = Total_Pendiente(noAgora, p.Key, listaEmpresas, listaSegmentos, "04", null);
                                workSheet.Cells[f, c].Value = totales_noagora_04;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;


                                //PARTE 05
                                //f++;
                                tamañoanterior = f.ToString();

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "05" & cadena[2] == "False")
                                    {
                                        workSheet.Cells[f, 3].Value = cadena[1] + " " + cadena[3];
                                        workSheet.Cells[f, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        workSheet.Cells[f, c].Value = Total_Pendiente(noAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                if (Pinto == true)
                                {
                                    RangoInterno = "B" + tamañoanterior;
                                    tamaño = f;
                                    tamaño = tamaño - 1;

                                    if (tamaño >= Int32.Parse(tamañoanterior))
                                    {
                                        RangoInterno = RangoInterno + ":B" + tamaño.ToString();
                                        workSheet.Cells[RangoInterno].Value = "DC no generado MDS";
                                        workSheet.Cells[RangoInterno].Style.Font.Bold = true;
                                        workSheet.Cells[RangoInterno].Merge = true;
                                        workSheet.Cells[RangoInterno].Style.WrapText = true;
                                        workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    }
                                }

                                workSheet.Cells[f, 3].Value = "Total DC no generado MDS";
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                RangoPintoGris = "B" + f.ToString() + ":C" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                RangoPintoGris = "B" + f.ToString() + ":M" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaGris);
                                totales_noagora_05 = Total_Pendiente(noAgora, p.Key, listaEmpresas, listaSegmentos, "05", null);
                                workSheet.Cells[f, c].Value = totales_noagora_05;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;

                                workSheet.Cells[f, 3].Value = "Total No Ágora";
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Value = totales_noagora_01 + totales_noagora_02 + totales_noagora_03 + totales_noagora_04 + totales_noagora_05;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                if (Pinto == true)
                                {
                                    
                                    RangoInterno = "A" + f.ToString() + ":M" + f.ToString();
                                    workSheet.Cells[RangoInterno].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    workSheet.Cells[RangoInterno].Style.Fill.BackgroundColor.SetColor(colorDeCeldaTotales);
                                    workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                }

                                f++;

                                // FIN PARTE 5

                                //f++;
                                tamañoanterior = f.ToString();
                                Pinto = false;

                                int o;
                                if (!dic_Totales_cups.TryGetValue(p.Key, out o))
                                    dic_Totales_cups.Add(p.Key, totales_noagora_01 + totales_noagora_02 + totales_noagora_03 + totales_noagora_04 + totales_noagora_05);

                            } //Fin if (dia < 6)
                        }

                        c = 14;

                        workSheet.Cells[3, 9].Value = "Pendiente Económico (€)";
                        workSheet.Cells[3, 9].Style.Font.Bold = true;
                        workSheet.Cells["I3:M3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //colorDeCelda = ColorTranslator.FromHtml("#c1dce6");          
                        workSheet.Cells["I3:M3"].Style.Fill.BackgroundColor.SetColor(colorDeCeldaCabecera);
                        workSheet.Cells["I3:M3"].Merge = true;
                        workSheet.Cells["I3:M3"].Style.WrapText = true;
                        workSheet.Cells["I3:M3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        workSheet.Cells["I3:M3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        // NO ÁGORA TAM ES PARTE ECONÓMICO*************************************
                        foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                        {
                            Console.WriteLine("Totales TAM ES noAgora dia: " + p.Key.ToString("dd/MM/yyyy"));
                            dia_tam++;

                            if (dia_tam < 6)
                            {
                                f = 4;
                                c--;

                                workSheet.Cells[f, c].Value = p.Key;
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                workSheet.Cells[f, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                //colorDeCelda = ColorTranslator.FromHtml("#5facad");
                                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(colorDeCeldaCabecera);
                                workSheet.Cells[f,c].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                f++;

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "01" & cadena[2] == "False")
                                    {
                                        //workSheet.Cells[f, 3].Value = cadena[1];
                                        workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }
                                totales_noagora_tam_01 = Total_Pendiente_TAM(noAgora, p.Key, listaEmpresas, listaSegmentos, "01", null);

                                workSheet.Cells[f, c].Value = totales_noagora_tam_01;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "02" & cadena[2] == "False")
                                    {
                                        //workSheet.Cells[f, 3].Value = cadena[1];
                                        workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                totales_noagora_tam_02 = Total_Pendiente_TAM(noAgora, p.Key, listaEmpresas, listaSegmentos, "02", null);

                                workSheet.Cells[f, c].Value = totales_noagora_tam_02;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "03" & cadena[2] == "False")
                                    {
                                        //workSheet.Cells[f, 3].Value = cadena[1];
                                        workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                totales_noagora_tam_03 = Total_Pendiente_TAM(noAgora, p.Key, listaEmpresas, listaSegmentos, "03", null);

                                workSheet.Cells[f, c].Value = totales_noagora_tam_03;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "04" & cadena[2] == "False")
                                    {
                                        //workSheet.Cells[f, 3].Value = cadena[1];
                                        workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }
                                totales_noagora_tam_04 = Total_Pendiente_TAM(noAgora, p.Key, listaEmpresas, listaSegmentos, "04", null);

                                workSheet.Cells[f, c].Value = totales_noagora_tam_04;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                                /*
                                if (blnPintoEtiqueta == true)
                                {
                                    // pego la etíqueta NO ÁGORA
                                    RangoInterno = "A5:A" + f.ToString();

                                    workSheet.Cells[RangoInterno].Value = "NO ÁGORA";
                                    workSheet.Cells[RangoInterno].Style.Font.Bold = true;
                                    workSheet.Cells[RangoInterno].Merge = true;
                                    workSheet.Cells[RangoInterno].Style.WrapText = true;
                                    workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                    blnPintoEtiqueta = false;
                                }

                                */

                                f++;

                                //PARTE 5
                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "05" & cadena[2] == "False")
                                    {
                                        //workSheet.Cells[f, 3].Value = cadena[1];
                                        workSheet.Cells[f, c].Value = Total_Pendiente_TAM(noAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }
                                totales_noagora_tam_05 = Total_Pendiente_TAM(noAgora, p.Key, listaEmpresas, listaSegmentos, "05", null);

                                workSheet.Cells[f, c].Value = totales_noagora_tam_05;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                if (blnPintoEtiqueta == true)
                                {
                                    // pego la etíqueta NO ÁGORA
                                    RangoInterno = "A5:A" + f.ToString();

                                    workSheet.Cells[RangoInterno].Value = "NO ÁGORA";
                                    workSheet.Cells[RangoInterno].Style.Font.Bold = true;
                                    workSheet.Cells[RangoInterno].Merge = true;
                                    workSheet.Cells[RangoInterno].Style.WrapText = true;
                                    workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                    workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    blnPintoEtiqueta = false;
                                }

                                f++;

                                //

                                workSheet.Cells[f, c].Value = totales_noagora_tam_01 + totales_noagora_tam_02 + totales_noagora_tam_03 + totales_noagora_tam_04 + totales_noagora_tam_05;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;

                                double o;
                                if (!dic_Totales_tam.TryGetValue(p.Key, out o))
                                    dic_Totales_tam.Add(p.Key, totales_noagora_tam_01 + totales_noagora_tam_02 + totales_noagora_tam_03 + totales_noagora_tam_04 + totales_noagora_tam_05);

                            }
                        }

                        c = 9;
                        dia = 0;
                        f--;
                        InicioRango = f--;

                        int tamañoprimerbloque;
                        tamañoprimerbloque = InicioRango;
                        Pinto = true;
                        blnPintoEtiqueta = true;

                        // ÁGORA ES  BLOQUE PENDIENTE **********************************************
                        foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                        {

                            dia++;

                            if (dia < 6)
                            {

                                f = InicioRango;
                                c--;
                                f++;

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "01" & cadena[2] == "True")
                                    {
                                        workSheet.Cells[f, 3].Value = cadena[1] + " " + cadena[3];
                                        workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                if (Pinto == true)
                                {
                                    tamaño = f;
                                    tamaño = tamaño - 1;
                                    tamañoprimerbloque++;
                                    if (tamaño >= tamañoprimerbloque) {
                                        RangoInterno = "B" + tamañoprimerbloque.ToString() + ":B" + tamaño.ToString();
                                        if (tamaño >= 1) //TOCADO
                                        {
                                        workSheet.Cells[RangoInterno].Value = "Pendiente medida";
                                        workSheet.Cells[RangoInterno].Style.Font.Bold = true;
                                        workSheet.Cells[RangoInterno].Merge = true;
                                        workSheet.Cells[RangoInterno].Style.WrapText = true;
                                        workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        }
                                    }
                                }

                                workSheet.Cells[f, 3].Value = "TotalPendiente medida";
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                RangoPintoGris = "B" + f.ToString() + ":C" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                RangoPintoGris = "B" + f.ToString() + ":M" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaGris);


                                /*
                                if (Pinto == true)
                                {
                                    colorDeCelda = ColorTranslator.FromHtml("#f0f4f5"); //Color GRIS
                                    RangoInterno = "B" + f.ToString() + ":M" + f.ToString();
                                    workSheet.Cells[RangoInterno].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    workSheet.Cells[RangoInterno].Style.Fill.BackgroundColor.SetColor(colorDeCelda);
                                }
                                */
                                totales_agora_01 = Total_Pendiente(siAgora, p.Key, listaEmpresas, listaSegmentos, "01", null);
                                workSheet.Cells[f, c].Value = totales_agora_01;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                f++;
                                tamañoanterior = f.ToString();

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "02" & cadena[2] == "True")
                                    {
                                        workSheet.Cells[f, 3].Value = cadena[1] + " " + cadena[3]; ;
                                        workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                if (Pinto == true)
                                {
                                    RangoInterno = "B" + tamañoanterior;
                                    tamaño = f;
                                    tamaño = tamaño - 1;
  
                                    if (tamaño >= Int32.Parse(tamañoanterior)) 
                                    {
                                        RangoInterno = RangoInterno + ":B" + tamaño.ToString();
                                        workSheet.Cells[RangoInterno].Value = "Orden de Cálculo Calculable";
                                        workSheet.Cells[RangoInterno].Style.Font.Bold = true;
                                        workSheet.Cells[RangoInterno].Merge = true;
                                        workSheet.Cells[RangoInterno].Style.WrapText = true;
                                        workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    }
                                }

                                workSheet.Cells[f, 3].Value = "Total Orden de Cálculo Calculable";
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                RangoPintoGris = "B" + f.ToString() + ":C" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                RangoPintoGris = "B" + f.ToString() + ":M" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaGris);

                                totales_agora_02 = Total_Pendiente(siAgora, p.Key, listaEmpresas, listaSegmentos, "02", null);
                                workSheet.Cells[f, c].Value = totales_agora_02;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;
                                tamañoanterior = f.ToString();

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "03" & cadena[2] == "True")
                                    {
                                        workSheet.Cells[f, 3].Value = cadena[1] + " " + cadena[3]; 
                                        workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                if (Pinto == true)
                                {
                                    RangoInterno = "B" + tamañoanterior;
                                    tamaño = f;
                                    tamaño = tamaño - 1;
                                   
                                    if (tamaño >= Int32.Parse(tamañoanterior)) 
                                    {
                                        RangoInterno = RangoInterno + ":B" + tamaño.ToString();
                                        workSheet.Cells[RangoInterno].Value = "DC Generado sin DI";
                                        workSheet.Cells[RangoInterno].Style.Font.Bold = true;
                                        workSheet.Cells[RangoInterno].Merge = true;
                                        workSheet.Cells[RangoInterno].Style.WrapText = true;
                                        workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    }
                                }

                                workSheet.Cells[f, 3].Value = "Total DC Generado sin DI";
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                RangoPintoGris = "B" + f.ToString() + ":C" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                RangoPintoGris = "B" + f.ToString() + ":M" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaGris);

                                totales_agora_03 = Total_Pendiente(siAgora, p.Key, listaEmpresas, listaSegmentos, "03", null);
                                workSheet.Cells[f, c].Value = totales_agora_03;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;
                                tamañoanterior = f.ToString();

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "04" & cadena[2] == "True")
                                    {
                                        workSheet.Cells[f, 3].Value = cadena[1] + " " + cadena[3];
                                        workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                if (Pinto == true)
                                {
                                    RangoInterno = "B" + tamañoanterior;
                                    tamaño = f;
                                    tamaño = tamaño - 1;
                                    
                                    if (tamaño >= Int32.Parse(tamañoanterior)) 
                                    {
                                        RangoInterno = RangoInterno + ":B" + tamaño.ToString();
                                        workSheet.Cells[RangoInterno].Value = "Doc. Impresión Apartado";
                                        workSheet.Cells[RangoInterno].Style.Font.Bold = true;
                                        workSheet.Cells[RangoInterno].Merge = true;
                                        workSheet.Cells[RangoInterno].Style.WrapText = true;
                                        workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    }
                                }

                                workSheet.Cells[f, 3].Value = "Total Doc. Impresión Apartado";
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                RangoPintoGris = "B" + f.ToString() + ":C" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                RangoPintoGris = "B" + f.ToString() + ":M" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaGris);
                                totales_agora_04 = Total_Pendiente(siAgora, p.Key, listaEmpresas, listaSegmentos, "04", null);
                                workSheet.Cells[f, c].Value = totales_agora_04;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;

                                //PARTE 5
                                tamañoanterior = f.ToString();

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "05" & cadena[2] == "True")
                                    {
                                        workSheet.Cells[f, 3].Value = cadena[1] + " " + cadena[3];
                                        workSheet.Cells[f, c].Value = Total_Pendiente(siAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                if (Pinto == true)
                                {
                                    RangoInterno = "B" + tamañoanterior;
                                    tamaño = f;
                                    tamaño = tamaño - 1;

                                    if (tamaño >= Int32.Parse(tamañoanterior))
                                    {
                                        RangoInterno = RangoInterno + ":B" + tamaño.ToString();
                                        workSheet.Cells[RangoInterno].Value = "Total DC no generado MDS";
                                        workSheet.Cells[RangoInterno].Style.Font.Bold = true;
                                        workSheet.Cells[RangoInterno].Merge = true;
                                        workSheet.Cells[RangoInterno].Style.WrapText = true;
                                        workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    }
                                }

                                workSheet.Cells[f, 3].Value = "Total DC no generado MDS";
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                RangoPintoGris = "B" + f.ToString() + ":C" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                RangoPintoGris = "B" + f.ToString() + ":M" + f.ToString();
                                workSheet.Cells[RangoPintoGris].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[RangoPintoGris].Style.Fill.BackgroundColor.SetColor(colorDeCeldaGris);
                                totales_agora_05 = Total_Pendiente(siAgora, p.Key, listaEmpresas, listaSegmentos, "05", null);
                                workSheet.Cells[f, c].Value = totales_agora_05;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;


                                //FIN PARTE 5

                                workSheet.Cells[f, 3].Value = "Total Si Ágora";
                                workSheet.Cells[f, 3].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Value = totales_agora_01 + totales_agora_02 + totales_agora_03 + totales_agora_04 + totales_agora_05;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                if (Pinto == true)
                                {
                                    //colorDeCelda = ColorTranslator.FromHtml("#e0f5c6"); //Color verde claro
                                    RangoInterno = "A" + f.ToString() + ":M" + f.ToString();
                                    workSheet.Cells[RangoInterno].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    workSheet.Cells[RangoInterno].Style.Fill.BackgroundColor.SetColor(colorDeCeldaTotales);
                                    workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                }

                                f++;

                                workSheet.Cells[f, 3].Value = "Total General";
                                
                                if (Pinto == true)
                                {
                                    //colorDeCelda = ColorTranslator.FromHtml("#c1dce6"); //Color azul claro
                                    RangoInterno = "A" + f.ToString() + ":M" + f.ToString();
                                    workSheet.Cells[RangoInterno].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    workSheet.Cells[RangoInterno].Style.Fill.BackgroundColor.SetColor(colorDeCeldaCabecera);
                                    workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                }
                                
                                //f++;

                                tamañoanterior = f.ToString();
                                Pinto = false;

                                int o;
                                if (dic_Totales_cups.TryGetValue(p.Key, out o))
                                {
                                    workSheet.Cells[f, c].Value = o + totales_agora_01 + totales_agora_02 + totales_agora_03 + totales_agora_04 + totales_agora_05;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, c].Style.Font.Bold = true;
                                    workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                }
                            }
                        }

                        c = 14;
                        dia_tam = 0;
                        pintoetiqueta = InicioRango;
                        pintoetiqueta++;

                        // ÁGORA TAM ES PARTE ECONÓMICA ************************************
                        foreach (KeyValuePair<DateTime, List<EndesaEntity.medida.Pendiente>> p in dic_pendiente_hist_fecha)
                        {

                            Console.WriteLine("Totales TAM ES siAgora dia: " + p.Key.ToString("dd/MM/yyyy"));

                            dia_tam++;

                            if (dia_tam < 6)
                            {
                                f = InicioRango;
                                c--;
                                f++;
                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "01" & cadena[2] == "True")
                                    {
                                        //workSheet.Cells[f, 3].Value = cadena[1];
                                        workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }
                                totales_agora_tam_01 = Total_Pendiente_TAM(siAgora, p.Key, listaEmpresas, listaSegmentos, "01", null);

                                workSheet.Cells[f, c].Value = totales_agora_tam_01;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "02" & cadena[2] == "True")
                                    {
                                        //workSheet.Cells[f, 3].Value = cadena[1];
                                        workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                totales_agora_tam_02 = Total_Pendiente_TAM(siAgora, p.Key, listaEmpresas, listaSegmentos, "02", null);

                                workSheet.Cells[f, c].Value = totales_agora_tam_02;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "03" & cadena[2] == "True")
                                    {
                                        //workSheet.Cells[f, 3].Value = cadena[1];
                                        workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }

                                totales_agora_tam_03 = Total_Pendiente_TAM(siAgora, p.Key, listaEmpresas, listaSegmentos, "03", null);

                                workSheet.Cells[f, c].Value = totales_agora_tam_03;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;

                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "04" & cadena[2] == "True")
                                    {
                                        //workSheet.Cells[f, 3].Value = cadena[1];
                                        workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }
                                totales_agora_tam_04 = Total_Pendiente_TAM(siAgora, p.Key, listaEmpresas, listaSegmentos, "04", null);

                                workSheet.Cells[f, c].Value = totales_agora_tam_04;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                f++;

                                //PARTE 5
                                for (int i = 1; i < B.Length; i++) //Quito Lenght-1 ya que en el último item están los nulos que no he rellenado al poner limite 5000
                                {
                                    string[] cadena = B[i].Split('_');
                                    if (cadena[0] == "05" & cadena[2] == "True")
                                    {
                                        //workSheet.Cells[f, 3].Value = cadena[1];
                                        workSheet.Cells[f, c].Value = Total_Pendiente_TAM(siAgora, p.Key, listaEmpresas, listaSegmentos, cadena[0], cadena[1]);
                                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                        workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        f++;
                                    }
                                }
                                totales_agora_tam_05 = Total_Pendiente_TAM(siAgora, p.Key, listaEmpresas, listaSegmentos, "05", null);

                                workSheet.Cells[f, c].Value = totales_agora_tam_05;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                // FIN PARTE 5

                                if (blnPintoEtiqueta == true)
                                {
                                    // pego la etíqueta SI ÁGORA
                                    RangoInterno = "A" + pintoetiqueta.ToString() + ":A" + f.ToString();

                                    workSheet.Cells[RangoInterno].Value = "SI ÁGORA";
                                    workSheet.Cells[RangoInterno].Style.Font.Bold = true;
                                    workSheet.Cells[RangoInterno].Merge = true;
                                    workSheet.Cells[RangoInterno].Style.WrapText = true;
                                    workSheet.Cells[RangoInterno].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                    workSheet.Cells[RangoInterno].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    blnPintoEtiqueta = false;

                                }

                                f++;

                                workSheet.Cells[f, c].Value = totales_agora_tam_01 + totales_agora_tam_02 + totales_agora_tam_03 + totales_agora_tam_04 + totales_agora_tam_05;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                workSheet.Cells[f, c].Style.Font.Bold = true;
                                workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                f++;

                                double o;
                                if (dic_Totales_tam.TryGetValue(p.Key, out o))
                                {
                                    workSheet.Cells[f, c].Value = o + totales_agora_tam_01 + totales_agora_tam_02 + totales_agora_tam_03 + totales_agora_tam_04 + totales_agora_tam_05;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                                    workSheet.Cells[f, c].Style.Font.Bold = true;
                                    workSheet.Cells[f, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                }

                                RangoInterno = "A1" + ":M" + f.ToString();
                                workSheet.Cells[RangoInterno].AutoFitColumns();
                                //El autofit parece no funcionar bien en las dos primeras columnas, las que están combinadas celdas, fuerzo el ancho
                                workSheet.Column(1).Width = 15;
                                workSheet.Column(2).Width = 20;
                            }

                            //allCells = workSheet.Cells[1, 1, f, 13];
                            //allCells.AutoFitColumns();
                            

                        }

                        //excelPackage.SaveAs(fileInfo);

                    } //Bucle HOJA
                    #endregion

                    #region Detalle ES

                    f = 1;
                    c = 1;

                    workSheet = excelPackage.Workbook.Worksheets.Add("Detalle ES");
                    headerCells = workSheet.Cells[1, 1, 1, 17];
                    headerFont = headerCells.Style.Font;

                    headerCells = workSheet.Cells[1, 1, 1, 30];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;
                    

                    workSheet.View.FreezePanes(2, 1);

                    DateTime fecha_informe = CalculaFechaDetalle();

                    strSql = "SELECT ps.cd_empr as EMPRESA,  ps.de_tp_cli as TIPO_CLIENTE, ps.cd_nif_cif_cli as NIF, ps.tx_apell_cli as CLIENTE,"
                        + " ps.fh_alta_crto as FALTACONT, ps.fh_inicio_vers_crto as FPSERCON, p.cd_cups as cups20, p.id_instalacion as N_INSTALACION, ps.cd_tarifa_c as TARIFA,"
                        + " ps.cd_crto_comercial as CONTRATO, p.fh_periodo as MES, ps.de_empr_distdora_nombre as DISTRIBUIDORA, "
                        + " concat(p.cd_estado,' ',de.de_estado) as ESTADO, concat(p.cd_subestado,' ',ds.de_subestado) as SUBESTADO, '' as DIAS_ESTADO , p.TAM as TAM, '' as MESES_PDTES_FACTURAR, '' as IMPORTE_PDTE_FACTURAR, IF(s.CUPS20 IS NULL, p.agora, 'S') AS agora,"
                        + " a.cd_tp_alarma AS TIPO_ALARMA, a.dcomenta AS ALARMA, AGR.nm_dia_agrup as DIA_AGRUPACION"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_d_alarmasconfact_pendiente_facturar a ON"
                        + " a.cd_cups_ext = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"
                        + " LEFT OUTER JOIN fact.cm_sofisticados s on"
                        + " s.CUPS20 = p.cd_cups"
                        + " LEFT OUTER JOIN "
                        + " (SELECT LEFT(cd_cups, 20) AS cups20, nm_dia_agrup FROM fact.SAP_AGRUPADAS_AGCUPS_New WHERE fh_carga = (select max(fh_carga) from fact.SAP_AGRUPADAS_AGCUPS_New) GROUP BY cd_cups) AS AGR"
                        + " ON AGR.cups20 = p.cd_cups"
                        + " where p.fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "' ";

                    //[26-02-2025] GUS: Añadimos códigos de empresas desde lista_empresas_ES en vez de hacerlo directamente [+ " p.cd_empr_titular in ('ES21','ES22')"];
                    firstOnly = true;

                    foreach (string p in lista_empresas_ES)
                    {
                        if (firstOnly)
                        {
                            strSql += " and p.cd_empr_titular in ("
                                + "'" + p + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + p + "'";
                    }

                    if(!firstOnly) //Si la lista tenía al menos un elemento cerramos sentencia correctamente.
                    {
                        strSql += ")";
                    }

                    

                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();

                    //Exporta cabeceras de las columnas que queremos sacar   - no nos vale en este caso, hay muchos 
                    //alias y calculos de columas
                    for (int i = 0; i < r.FieldCount; i++)
                    {
                        workSheet.Cells[1, i + 1].Value = r.GetName(i);
                    }

                    //Ponemos columnas directamente
                    /*
                    workSheet.Cells[1, 1].Value = "EMPRESA";
                    workSheet.Cells[1, 2].Value = "TIPO CLIENTE";
                    workSheet.Cells[1, 3].Value = "NIF";
                    workSheet.Cells[1, 4].Value = "CLIENTE";
                    workSheet.Cells[1, 5].Value = "FALTACONT";
                    workSheet.Cells[1, 6].Value = "FPSERCON";
                    workSheet.Cells[1, 7].Value = "CUPS20";
                    workSheet.Cells[1, 8].Value = "Nº INSTALACIÓN";
                    workSheet.Cells[1, 9].Value = "TARIFA";
                    workSheet.Cells[1, 10].Value = "CONTRATO";
                    workSheet.Cells[1, 11].Value = "MES";
                    workSheet.Cells[1, 12].Value = "DISTRIBUIDORA";
                    workSheet.Cells[1, 13].Value = "ESTADO";
                    workSheet.Cells[1, 14].Value = "SUBESTADO";
                    workSheet.Cells[1, 15].Value = "DIAS ESTADO";
                    workSheet.Cells[1, 16].Value = "TAM";
                    workSheet.Cells[1, 17].Value = "MESES PDTES FACTURAR";
                    workSheet.Cells[1, 18].Value = "IMPORTE PDTE FACTURAR";
                    workSheet.Cells[1, 19].Value = "ÁGORA";                 
                    workSheet.Cells[1, 20].Value = "TIPO ALARMA";    
                    workSheet.Cells[1, 21].Value = "ALARMA";    
                    workSheet.Cells[1, 22].Value = "DIA AGRUPACION";      
                    */

                    while (r.Read())
                    {
                        f++;
                        c = 1;

                        #region BUSCAR DATOS EN T_ED_H_PS_HIST
                        //Si el NIF está vacío supuestamente es porque el contrato está de baja y sale de la tabla cont.t_ed_h_ps
                        //Intentamos recuperar los campos EMPRESA, TIPO_CLIENTE, NIF, CLIENTE, TARIFA, CONTRATO Y DISTRIBUIDORA del histórico, cont.t_ed_h_ps_hist
                        Dictionary<string, string> datos_histo = new Dictionary<string, string>();
                        string dat = "";
                        if (r["NIF"] == System.DBNull.Value && r["CUPS20"] != System.DBNull.Value)
                        {
                            strSql_histo = "SELECT ps.cd_empr AS EMPRESA, ps.de_tp_cli AS TIPO_CLIENTE, ps.cd_nif_cif_cli AS NIF, ps.tx_apell_cli AS CLIENTE, ps.cd_tarifa_c AS TARIFA, ps.cd_crto_comercial AS CONTRATO, ps.de_empr_distdora_nombre AS DISTRIBUIDORA "
                                + "FROM cont.t_ed_h_ps_hist ps WHERE ps.cups20 = '"+ r["CUPS20"] + "' ORDER BY ps.fh_act_dmco DESC LIMIT 1;";
                            
                            db_histo = new MySQLDB(MySQLDB.Esquemas.GBL);
                            command_histo = new MySqlCommand(strSql_histo, db_histo.con);
                            r_histo = command_histo.ExecuteReader();
                            if (r_histo.Read())
                            {
                                if (r_histo["EMPRESA"] != System.DBNull.Value)
                                    datos_histo.Add("EMPRESA", r_histo["EMPRESA"].ToString());
                                if (r_histo["TIPO_CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("TIPO_CLIENTE", r_histo["TIPO_CLIENTE"].ToString());
                                if (r_histo["NIF"] != System.DBNull.Value)
                                    datos_histo.Add("NIF", r_histo["NIF"].ToString());
                                if (r_histo["CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("CLIENTE", r_histo["CLIENTE"].ToString());
                                if (r_histo["TARIFA"] != System.DBNull.Value)
                                    datos_histo.Add("TARIFA", r_histo["TARIFA"].ToString());
                                if (r_histo["CONTRATO"] != System.DBNull.Value)
                                    datos_histo.Add("CONTRATO", r_histo["CONTRATO"].ToString());
                                if (r_histo["DISTRIBUIDORA"] != System.DBNull.Value)
                                    datos_histo.Add("DISTRIBUIDORA", r_histo["DISTRIBUIDORA"].ToString());
                            }

                            db_histo.CloseConnection();
                        }
                        #endregion

                        //Creamos objeto ResumenAgrupadaPendiente para ir almacenando info y al final, si procede, añadir a dic dic_resumen_agrupadas_ES
                        EndesaEntity.facturacion.ResumenAgrupadaPendiente rap = new EndesaEntity.facturacion.ResumenAgrupadaPendiente();
                        tiene_num_dia_agrupacion = false;
                        

                        if (r["EMPRESA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["EMPRESA"].ToString();
                        else if (datos_histo.TryGetValue("EMPRESA",out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["TIPO_CLIENTE"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TIPO_CLIENTE"].ToString();
                        else if (datos_histo.TryGetValue("TIPO_CLIENTE", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["NIF"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["NIF"].ToString();
                            rap.cif = r["NIF"].ToString();
                        }
                        else if (datos_histo.TryGetValue("NIF", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            rap.cif = dat;
                        }
                        c++;

                        if (r["CLIENTE"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["CLIENTE"].ToString();
                            rap.razon_social = r["CLIENTE"].ToString();
                        }
                        else if (datos_histo.TryGetValue("CLIENTE", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            rap.razon_social = dat;
                        }
                        c++;

                        if (r["FALTACONT"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["FALTACONT"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["FPSERCON"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["FPSERCON"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["CUPS20"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CUPS20"].ToString();
                        c++;

                        if (r["N_INSTALACION"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["N_INSTALACION"].ToString();
                        c++;

                        if (r["TARIFA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                        else if (datos_histo.TryGetValue("TARIFA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["CONTRATO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CONTRATO"].ToString();
                        else if (datos_histo.TryGetValue("CONTRATO", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["MES"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["MES"]);
                            aniomes = Convert.ToInt32(r["MES"]);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                        }
                        c++;

                        if (r["DISTRIBUIDORA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["DISTRIBUIDORA"].ToString();
                        else if (datos_histo.TryGetValue("DISTRIBUIDORA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["ESTADO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["ESTADO"].ToString();
                        c++;

                        if (r["SUBESTADO"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["SUBESTADO"].ToString();
                            switch(r["SUBESTADO"].ToString().Substring(0,2))
                            {
                                case "01":
                                    rap.pendiente_medida = 1;
                                    break;
                                case "02":
                                    rap.oc_calculable = 1;
                                    break;
                                case "03":
                                    rap.dc_generado_sin_di = 1;
                                    break;
                                case "04":
                                    rap.di_apartado = 1;
                                    break;

                            }
                        }
                        c++;

                        //if (r["lg_multimedida"] != System.DBNull.Value)
                        //{
                        //    workSheet.Cells[f, c].Value = r["lg_multimedida"].ToString();

                        //}
                        //else
                        //{
                        //    workSheet.Cells[f, c].Value = "N";
                        //}

                        workSheet.Cells[f, c].Value = GetDiasEstado(r["CUPS20"].ToString());

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++; 

                        if (r["TAM"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            rap.tam_total = rap.tam_total + Convert.ToDouble(r["TAM"]);
                        }
                        c++;

                        if (r["MES"] != System.DBNull.Value)
                        {
                            //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                            workSheet.Cells[f, c].Value = meses_pdtes;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        c++;

                        if (r["MES"] != System.DBNull.Value && r["TAM"] != System.DBNull.Value)
                        {
                            //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;                            
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]) * meses_pdtes;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        }
                        else
                        {
                            c++;
                        }

                        if (r["AGORA"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["AGORA"].ToString();
                            if (r["AGORA"].ToString() == "S")
                                rap.es_agora = true;
                        }

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["TIPO_ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TIPO_ALARMA"].ToString();

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["ALARMA"].ToString();

                        c++;

                        if (r["DIA_AGRUPACION"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["DIA_AGRUPACION"].ToString();
                            rap.num_dia_agrupacion = Convert.ToInt16(r["DIA_AGRUPACION"]);
                            tiene_num_dia_agrupacion = true;
                        }

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        //Si tiene dia de agrupación, buscamos en dic y modificamos o añadimos si no existe aún ese NIF/CIF en dic_resumen_agrupadas_ES
                        if (tiene_num_dia_agrupacion && (rap.cif != null))
                        {
                            EndesaEntity.facturacion.ResumenAgrupadaPendiente o_rap = new EndesaEntity.facturacion.ResumenAgrupadaPendiente();
                            if (!dic_resumen_agrupadas_ES.TryGetValue(rap.cif, out o_rap))
                            {
                                dic_resumen_agrupadas_ES.Add(rap.cif, rap);
                            }
                            else
                            {
                                o_rap.pendiente_medida += rap.pendiente_medida;
                                o_rap.oc_calculable += rap.oc_calculable;
                                o_rap.dc_generado_sin_di += rap.dc_generado_sin_di;
                                o_rap.di_apartado += rap.di_apartado;
                                o_rap.tam_total += rap.tam_total;
                                if (rap.es_agora)
                                    o_rap.es_agora = true;
                            }
                        }
                    }
                    db.CloseConnection();

                    headerCells = workSheet.Cells[1, 1, 1, c];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;
                    allCells = workSheet.Cells[1, 1, f, c];

                    allCells.AutoFitColumns();

                    //workSheet.View.FreezePanes(2, 1);
                    //workSheet.Cells["A1:S1"].AutoFilter = true;
                    //allCells.AutoFitColumns();

                    #endregion

                    #region Detalle POR MT-BTE
                    f = 1;
                    c = 1;

                    workSheet = excelPackage.Workbook.Worksheets.Add("Detalle POR MT-BTE");
                    headerCells = workSheet.Cells[1, 1, 1, 17];
                    headerFont = headerCells.Style.Font;

                    headerCells = workSheet.Cells[1, 1, 1, 30];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;

                    workSheet.View.FreezePanes(2, 1);

                    allCells = workSheet.Cells[1, 1, 50, 50];

                    strSql = "SELECT ps.cd_empr, p.cd_segmento_ptg, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                        + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, p.cd_cups as cups20, p.id_instalacion, ps.cd_tarifa_c,"
                        + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre,  p.lg_multimedida,"
                        + " concat(p.cd_estado,' ',de.de_estado) as de_estado, concat(p.cd_subestado,' ',ds.de_subestado) as de_subestado, p.fh_periodo as mes, IF(s.CUPS20 IS NULL, p.agora, 'S') AS agora, p.TAM,"
                        + " a.cd_tp_alarma AS TIPO_ALARMA, a.dcomenta AS ALARMA, AGR.nm_dia_agrup as DIA_AGRUPACION"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps_pt ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_d_alarmasconfact_pendiente_facturar a ON"
                        + " a.cd_cups_ext = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"
                        + " LEFT OUTER JOIN fact.cm_sofisticados s ON"
                        + " s.CUPS20 = p.cd_cups"
                        + " LEFT OUTER JOIN "
                        + " (SELECT LEFT(cd_cups, 20) AS cups20, nm_dia_agrup FROM fact.SAP_AGRUPADAS_AGCUPS_New WHERE fh_carga = (select max(fh_carga) from fact.SAP_AGRUPADAS_AGCUPS_New) GROUP BY cd_cups) AS AGR"
                        + " ON AGR.cups20 = p.cd_cups"
                        + " where p.fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "' ";
                    
                    //[26-02-2025] GUS: Añadimos códigos de empresas desde lista_empresas_ES
                    // y segmentos de Portugal desde lista_segmentos_MT_BTE
                    // en vez de hacerlo directamente [+ " p.cd_empr_titular = 'PT1Q' and p.cd_segmento_ptg in ('MT','BTE')";]

                    firstOnly = true;
                    foreach (string p in lista_empresas_PT)
                    {
                        if (firstOnly)
                        {
                            strSql += " and p.cd_empr_titular in ("
                                + "'" + p + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + p + "'";
                    }
                    if (!firstOnly) //Si la lista tenía al menos un elemento cerramos sentencia correctamente.
                    {
                        strSql += ")";
                    }


                    firstOnly = true;
                    foreach (string p in lista_segmentos_MT_BTE)
                    {
                        if (firstOnly)
                        {
                            strSql += " and p.cd_segmento_ptg in ("
                                + "'" + p + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + p + "'";
                    }
                    if (!firstOnly) //Si la lista tenía al menos un elemento cerramos sentencia correctamente.
                    {
                        strSql += ")";
                    }

                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();


                    //Ponemos columnas directamente
                    workSheet.Cells[1, 1].Value = "EMPRESA";
                    workSheet.Cells[1, 2].Value = "SEGMENTO";
                    workSheet.Cells[1, 3].Value = "TIPO CLIENTE";
                    workSheet.Cells[1, 4].Value = "NIF";
                    workSheet.Cells[1, 5].Value = "CLIENTE";
                    workSheet.Cells[1, 6].Value = "FALTACONT";
                    workSheet.Cells[1, 7].Value = "FPSERCON";
                    workSheet.Cells[1, 8].Value = "CUPS20";
                    workSheet.Cells[1, 9].Value = "Nº INSTALACIÓN";
                    workSheet.Cells[1, 10].Value = "TARIFA";
                    workSheet.Cells[1, 11].Value = "CONTRATO";
                    workSheet.Cells[1, 12].Value = "MES";
                    workSheet.Cells[1, 13].Value = "DISTRIBUIDORA";
                    workSheet.Cells[1, 14].Value = "ESTADO";
                    workSheet.Cells[1, 15].Value = "SUBESTADO";
                    workSheet.Cells[1, 16].Value = "DIAS ESTADO";
                    workSheet.Cells[1, 17].Value = "TAM";
                    workSheet.Cells[1, 18].Value = "MESES PDTES FACTURAR";
                    workSheet.Cells[1, 19].Value = "IMPORTE PDTE FACTURAR";
                    workSheet.Cells[1, 20].Value = "ÁGORA";
                    workSheet.Cells[1, 21].Value = "TIPO ALARMA";
                    workSheet.Cells[1, 22].Value = "ALARMA";
                    workSheet.Cells[1, 23].Value = "DIA AGRUPACION";

                    while (r.Read())
                    {
                        f++;
                        c = 1;

                        #region BUSCAR DATOS EN T_ED_H_PS_PT_HIST (MT-BTE)
                        //Si el NIF está vacío supuestamente es porque el contrato está de baja y sale de la tabla cont.t_ed_h_ps_pt
                        //Intentamos recuperar los campos EMPRESA, TIPO_CLIENTE, NIF, CLIENTE, TARIFA, CONTRATO Y DISTRIBUIDORA del histórico, cont.t_ed_h_ps_pt_hist
                        Dictionary<string, string> datos_histo = new Dictionary<string, string>();
                        string dat = "";
                        if (r["cd_nif_cif_cli"] == System.DBNull.Value && r["cups20"] != System.DBNull.Value)
                        {
                            strSql_histo = "SELECT ps.cd_empr AS EMPRESA, ps.de_tp_cli AS TIPO_CLIENTE, ps.cd_nif_cif_cli AS NIF, ps.tx_apell_cli AS CLIENTE, ps.cd_tarifa_c AS TARIFA, ps.cd_crto_comercial AS CONTRATO, ps.de_empr_distdora_nombre AS DISTRIBUIDORA "
                                + "FROM cont.t_ed_h_ps_pt_hist ps WHERE ps.cups20 = '" + r["cups20"] + "' ORDER BY ps.fh_act_dmco DESC LIMIT 1;";

                            db_histo = new MySQLDB(MySQLDB.Esquemas.GBL);
                            command_histo = new MySqlCommand(strSql_histo, db_histo.con);
                            r_histo = command_histo.ExecuteReader();
                           // r_histo.Read();

                            while (r_histo.Read())
                            {
                                if (r_histo["EMPRESA"] != System.DBNull.Value)
                                    datos_histo.Add("EMPRESA", r_histo["EMPRESA"].ToString());
                                if (r_histo["TIPO_CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("TIPO_CLIENTE", r_histo["TIPO_CLIENTE"].ToString());
                                if (r_histo["NIF"] != System.DBNull.Value)
                                    datos_histo.Add("NIF", r_histo["NIF"].ToString());
                                if (r_histo["CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("CLIENTE", r_histo["CLIENTE"].ToString());
                                if (r_histo["TARIFA"] != System.DBNull.Value)
                                    datos_histo.Add("TARIFA", r_histo["TARIFA"].ToString());
                                if (r_histo["CONTRATO"] != System.DBNull.Value)
                                    datos_histo.Add("CONTRATO", r_histo["CONTRATO"].ToString());
                                if (r_histo["DISTRIBUIDORA"] != System.DBNull.Value)
                                    datos_histo.Add("DISTRIBUIDORA", r_histo["DISTRIBUIDORA"].ToString());
                            }

                            db_histo.CloseConnection();
                        }
                        #endregion

                        //Creamos objeto ResumenAgrupadaPendiente para ir almacenando info y al final, si procede, añadir a dic dic_resumen_agrupadas_MTBTE
                        EndesaEntity.facturacion.ResumenAgrupadaPendiente rap = new EndesaEntity.facturacion.ResumenAgrupadaPendiente();
                        tiene_num_dia_agrupacion = false;


                        if (r["cd_empr"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cd_empr"].ToString();
                        else if (datos_histo.TryGetValue("EMPRESA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["cd_segmento_ptg"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cd_segmento_ptg"].ToString();
                        else
                            workSheet.Cells[f, c].Value = "NULO";
                        c++;

                        if (r["de_tp_cli"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["de_tp_cli"].ToString();
                        else if (datos_histo.TryGetValue("TIPO_CLIENTE", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["cd_nif_cif_cli"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["cd_nif_cif_cli"].ToString();
                            rap.cif = r["cd_nif_cif_cli"].ToString();
                        }
                        else if (datos_histo.TryGetValue("NIF", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            rap.cif = dat;
                        }
                            c++;

                        if (r["tx_apell_cli"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["tx_apell_cli"].ToString();
                            rap.razon_social = r["tx_apell_cli"].ToString();
                        }
                        else if (datos_histo.TryGetValue("CLIENTE", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            rap.razon_social = dat;
                        }   
                        c++;

                        if (r["fh_alta_crto"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_alta_crto"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["fh_inicio_vers_crto"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_inicio_vers_crto"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["cups20"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cups20"].ToString();
                        c++;

                        if (r["id_instalacion"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["id_instalacion"].ToString();
                        c++;

                        if (r["cd_tarifa_c"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cd_tarifa_c"].ToString();
                        else if (datos_histo.TryGetValue("TARIFA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["cd_crto_comercial"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cd_crto_comercial"].ToString();
                        else if (datos_histo.TryGetValue("CONTRATO", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["mes"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["mes"]);
                            aniomes = Convert.ToInt32(r["mes"]);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                        }
                        c++;

                        if (r["de_empr_distdora_nombre"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["de_empr_distdora_nombre"].ToString();
                        else if (datos_histo.TryGetValue("DISTRIBUIDORA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["de_estado"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["de_estado"].ToString();
                        c++;

                        if (r["de_subestado"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["de_subestado"].ToString();
                            switch (r["de_subestado"].ToString().Substring(0, 2))
                            {
                                case "01":
                                    rap.pendiente_medida = 1;
                                    break;
                                case "02":
                                    rap.oc_calculable = 1;
                                    break;
                                case "03":
                                    rap.dc_generado_sin_di = 1;
                                    break;
                                case "04":
                                    rap.di_apartado = 1;
                                    break;

                            }
                        }
                        c++;

                        workSheet.Cells[f, c].Value = GetDiasEstado(r["cups20"].ToString());

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["TAM"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            rap.tam_total = rap.tam_total + Convert.ToDouble(r["TAM"]);
                        }
                        c++;

                        if (r["mes"] != System.DBNull.Value)
                        {
                            //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                            workSheet.Cells[f, c].Value = meses_pdtes;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
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

                        if (r["agora"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["agora"].ToString();
                            if (r["agora"].ToString() == "S")
                                rap.es_agora = true;
                        }

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["TIPO_ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TIPO_ALARMA"].ToString();

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["ALARMA"].ToString();

                        c++;

                        if (r["DIA_AGRUPACION"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["DIA_AGRUPACION"].ToString();
                            rap.num_dia_agrupacion = Convert.ToInt16(r["DIA_AGRUPACION"]);
                            tiene_num_dia_agrupacion = true;
                        }

                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        //Si tiene dia de agrupación, buscamos en dic y modificamos o añadimos si no existe aún ese NIF/CIF en dic_resumen_agrupadas_MTBTE
                        if (tiene_num_dia_agrupacion && (rap.cif != null))
                        {
                            EndesaEntity.facturacion.ResumenAgrupadaPendiente o_rap = new EndesaEntity.facturacion.ResumenAgrupadaPendiente();
                            if (!dic_resumen_agrupadas_MTBTE .TryGetValue(rap.cif, out o_rap))
                            {
                                dic_resumen_agrupadas_MTBTE.Add(rap.cif, rap);
                            }
                            else
                            {
                                o_rap.pendiente_medida += rap.pendiente_medida;
                                o_rap.oc_calculable += rap.oc_calculable;
                                o_rap.dc_generado_sin_di += rap.dc_generado_sin_di;
                                o_rap.di_apartado += rap.di_apartado;
                                o_rap.tam_total += rap.tam_total;
                                if (rap.es_agora)
                                    o_rap.es_agora = true;
                            }
                        }
                    }
                    db.CloseConnection();

                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    #endregion

                    #region Detalle POR BTN

                    f = 1;
                    c = 1;

                    workSheet = excelPackage.Workbook.Worksheets.Add("Detalle POR BTN");
                    headerCells = workSheet.Cells[1, 1, 1, 17];
                    headerFont = headerCells.Style.Font;

                    headerCells = workSheet.Cells[1, 1, 1, 30];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;
                    
                    workSheet.View.FreezePanes(2, 1);

                    allCells = workSheet.Cells[1, 1, 50, 50];

                    strSql = "SELECT ps.cd_empr, p.cd_segmento_ptg, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                        + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, p.cd_cups as cups20, p.id_instalacion, ps.cd_tarifa_c,"
                        + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre,  p.lg_multimedida,"
                        + " concat(p.cd_estado,' ',de.de_estado) as de_estado, concat(p.cd_subestado,' ',ds.de_subestado) as de_subestado, p.fh_periodo as mes, IF(s.CUPS20 IS NULL, p.agora, 'S') AS agora, p.TAM,"
                        + " a.cd_tp_alarma AS TIPO_ALARMA, a.dcomenta AS ALARMA, AGR.nm_dia_agrup as DIA_AGRUPACION, DATE_SUB(p.fh_desde, INTERVAL 1 DAY) as fh_hasta_ultima_factura"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps_pt ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_d_alarmasconfact_pendiente_facturar a ON"
                        + " a.cd_cups_ext = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"
                        + " LEFT OUTER JOIN fact.cm_sofisticados s ON"
                        + " s.CUPS20 = p.cd_cups "
                        + " LEFT OUTER JOIN "
                        + " (SELECT LEFT(cd_cups, 20) AS cups20, nm_dia_agrup FROM fact.SAP_AGRUPADAS_AGCUPS_New WHERE fh_carga = (select max(fh_carga) from fact.SAP_AGRUPADAS_AGCUPS_New) GROUP BY cd_cups) AS AGR"
                        + " ON AGR.cups20 = p.cd_cups"
                        + " where p.fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "' ";
                        
                    //[26-02-2025] GUS: Añadimos códigos de empresas desde lista_empresas_ES
                    // y segmentos de Portugal desde lista_segmentos_BTN
                    // en vez de hacerlo directamente [+ " p.cd_empr_titular = 'PT1Q' and p.cd_segmento_ptg = 'BTN'";]
                    // Además incluimos aquellos registros que no tienen informado el segmento, lo tienen a null (a petición de Ignacio)

                    firstOnly = true;
                    foreach (string p in lista_empresas_PT)
                    {
                        if (firstOnly)
                        {
                            strSql += " and p.cd_empr_titular in ("
                                + "'" + p + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + p + "'";
                    }
                    if (!firstOnly) //Si la lista tenía al menos un elemento cerramos sentencia correctamente.
                    {
                        strSql += ")";
                    }

                    firstOnly = true;
                    foreach (string p in lista_segmentos_BTN)
                    {
                        if (firstOnly)
                        {
                            strSql += " and (p.cd_segmento_ptg in ("
                                + "'" + p + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + p + "'";
                    }
                    if (!firstOnly) //Si la lista tenía al menos un elemento cerramos sentencia correctamente.
                    {
                        strSql += ") or p.cd_segmento_ptg is null)";
                    }
                    else
                        strSql += " and p.cd_segmento_ptg is null";

                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();

                    //Ponemos columnas directamente
                    workSheet.Cells[1, 1].Value = "EMPRESA";
                    workSheet.Cells[1, 2].Value = "SEGMENTO";
                    workSheet.Cells[1, 3].Value = "TIPO CLIENTE";
                    workSheet.Cells[1, 4].Value = "NIF";
                    workSheet.Cells[1, 5].Value = "CLIENTE";
                    workSheet.Cells[1, 6].Value = "FALTACONT";
                    workSheet.Cells[1, 7].Value = "FPSERCON";
                    workSheet.Cells[1, 8].Value = "CUPS20";
                    workSheet.Cells[1, 9].Value = "Nº INSTALACIÓN";
                    workSheet.Cells[1, 10].Value = "TARIFA";
                    workSheet.Cells[1, 11].Value = "CONTRATO";
                    workSheet.Cells[1, 12].Value = "MES";
                    workSheet.Cells[1, 13].Value = "DISTRIBUIDORA";
                    workSheet.Cells[1, 14].Value = "ESTADO";
                    workSheet.Cells[1, 15].Value = "SUBESTADO";
                    workSheet.Cells[1, 16].Value = "DIAS ESTADO";
                    workSheet.Cells[1, 17].Value = "TAM";
                    workSheet.Cells[1, 18].Value = "MESES PDTES FACTURAR";
                    workSheet.Cells[1, 19].Value = "IMPORTE PDTE FACTURAR";
                    workSheet.Cells[1, 20].Value = "ÁGORA";
                    workSheet.Cells[1, 21].Value = "TIPO ALARMA";
                    workSheet.Cells[1, 22].Value = "ALARMA";
                    workSheet.Cells[1, 23].Value = "DIA AGRUPACION";
                    workSheet.Cells[1, 24].Value = "ULTIMA LECTURA FACTURADA";

                    while (r.Read())
                    {
                        f++;
                        c = 1;


                        #region BUSCAR DATOS EN T_ED_H_PS_PT_HIST (BTN)
                        //Si el NIF está vacío supuestamente es porque el contrato está de baja y sale de la tabla cont.t_ed_h_ps_pt
                        //Intentamos recuperar los campos EMPRESA, TIPO_CLIENTE, NIF, CLIENTE, TARIFA, CONTRATO Y DISTRIBUIDORA del histórico, cont.t_ed_h_ps_pt_hist
                        Dictionary<string, string> datos_histo = new Dictionary<string, string>();
                        string dat = "";
                        if (r["cd_nif_cif_cli"] == System.DBNull.Value && r["cups20"] != System.DBNull.Value)
                        {
                            strSql_histo = "SELECT ps.cd_empr AS EMPRESA, ps.de_tp_cli AS TIPO_CLIENTE, ps.cd_nif_cif_cli AS NIF, ps.tx_apell_cli AS CLIENTE, ps.cd_tarifa_c AS TARIFA, ps.cd_crto_comercial AS CONTRATO, ps.de_empr_distdora_nombre AS DISTRIBUIDORA "
                                + "FROM cont.t_ed_h_ps_pt_hist ps WHERE ps.cups20 = '" + r["cups20"] + "' ORDER BY ps.fh_act_dmco DESC LIMIT 1;";

                            db_histo = new MySQLDB(MySQLDB.Esquemas.GBL);
                            command_histo = new MySqlCommand(strSql_histo, db_histo.con);
                            r_histo = command_histo.ExecuteReader();
                            r_histo.Read();

                            while (r_histo.Read())
                            {
                                if (r_histo["EMPRESA"] != System.DBNull.Value)
                                    datos_histo.Add("EMPRESA", r_histo["EMPRESA"].ToString());
                                if (r_histo["TIPO_CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("TIPO_CLIENTE", r_histo["TIPO_CLIENTE"].ToString());
                                if (r_histo["NIF"] != System.DBNull.Value)
                                    datos_histo.Add("NIF", r_histo["NIF"].ToString());
                                if (r_histo["CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("CLIENTE", r_histo["CLIENTE"].ToString());
                                if (r_histo["TARIFA"] != System.DBNull.Value)
                                    datos_histo.Add("TARIFA", r_histo["TARIFA"].ToString());
                                if (r_histo["CONTRATO"] != System.DBNull.Value)
                                    datos_histo.Add("CONTRATO", r_histo["CONTRATO"].ToString());
                                if (r_histo["DISTRIBUIDORA"] != System.DBNull.Value)
                                    datos_histo.Add("DISTRIBUIDORA", r_histo["DISTRIBUIDORA"].ToString());
                            }

                            db_histo.CloseConnection();
                        }
                        #endregion

                        //Creamos objeto ResumenAgrupadaPendiente para ir almacenando info y al final, si procede, añadir a dic dic_resumen_agrupadas_BTN
                        EndesaEntity.facturacion.ResumenAgrupadaPendiente rap = new EndesaEntity.facturacion.ResumenAgrupadaPendiente();
                        tiene_num_dia_agrupacion = false;

                        if (r["cd_empr"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cd_empr"].ToString();
                        else if (datos_histo.TryGetValue("EMPRESA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["cd_segmento_ptg"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cd_segmento_ptg"].ToString();
                        else
                            workSheet.Cells[f, c].Value = "NULO";
                        c++;

                        if (r["de_tp_cli"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["de_tp_cli"].ToString();
                        else if (datos_histo.TryGetValue("TIPO_CLIENTE", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["cd_nif_cif_cli"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["cd_nif_cif_cli"].ToString();
                            rap.cif = r["cd_nif_cif_cli"].ToString();
                        }
                        else if (datos_histo.TryGetValue("NIF", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            rap.cif = dat;
                        }
                        c++;

                        if (r["tx_apell_cli"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["tx_apell_cli"].ToString();
                            rap.razon_social = r["tx_apell_cli"].ToString();
                        }
                        else if (datos_histo.TryGetValue("CLIENTE", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            rap.razon_social = dat;
                        }
                        c++;

                        if (r["fh_alta_crto"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_alta_crto"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["fh_inicio_vers_crto"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_inicio_vers_crto"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["cups20"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cups20"].ToString();
                        c++;

                        if (r["id_instalacion"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["id_instalacion"].ToString();
                        c++;

                        if (r["cd_tarifa_c"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cd_tarifa_c"].ToString();
                        else if (datos_histo.TryGetValue("TARIFA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["cd_crto_comercial"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cd_crto_comercial"].ToString();
                        else if (datos_histo.TryGetValue("CONTRATO", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["mes"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["mes"]);
                            aniomes = Convert.ToInt32(r["mes"]);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                        }
                        c++;

                        if (r["de_empr_distdora_nombre"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["de_empr_distdora_nombre"].ToString();
                        else if (datos_histo.TryGetValue("DISTRIBUIDORA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["de_estado"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["de_estado"].ToString();
                        c++;

                        if (r["de_subestado"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["de_subestado"].ToString();
                            switch (r["de_subestado"].ToString().Substring(0, 2))
                            {
                                case "01":
                                    rap.pendiente_medida = 1;
                                    break;
                                case "02":
                                    rap.oc_calculable = 1;
                                    break;
                                case "03":
                                    rap.dc_generado_sin_di = 1;
                                    break;
                                case "04":
                                    rap.di_apartado = 1;
                                    break;

                            }
                        }
                        c++;

                        workSheet.Cells[f, c].Value = GetDiasEstado(r["cups20"].ToString());
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["TAM"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            rap.tam_total = rap.tam_total + Convert.ToDouble(r["TAM"]);
                        }
                        c++;

                        if (r["mes"] != System.DBNull.Value)
                        {
                            //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                            workSheet.Cells[f, c].Value = meses_pdtes;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
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

                        if (r["agora"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["agora"].ToString();

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["TIPO_ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TIPO_ALARMA"].ToString();

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["ALARMA"].ToString();
                        c++;

                        if (r["DIA_AGRUPACION"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["DIA_AGRUPACION"].ToString();
                            rap.num_dia_agrupacion = Convert.ToInt16(r["DIA_AGRUPACION"]);
                            tiene_num_dia_agrupacion = true;
                        }
                        
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["fh_hasta_ultima_factura"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["fh_hasta_ultima_factura"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        //Si tiene dia de agrupación, buscamos en dic y modificamos o añadimos si no existe aún ese NIF/CIF en dic_resumen_agrupadas_BTN
                        if (tiene_num_dia_agrupacion && (rap.cif != null))
                        {
                            EndesaEntity.facturacion.ResumenAgrupadaPendiente o_rap = new EndesaEntity.facturacion.ResumenAgrupadaPendiente();
                            if (!dic_resumen_agrupadas_BTN.TryGetValue(rap.cif, out o_rap))
                            {
                                dic_resumen_agrupadas_BTN.Add(rap.cif, rap);
                            }
                            else
                            {
                                o_rap.pendiente_medida += rap.pendiente_medida;
                                o_rap.oc_calculable += rap.oc_calculable;
                                o_rap.dc_generado_sin_di += rap.dc_generado_sin_di;
                                o_rap.di_apartado += rap.di_apartado;
                                o_rap.tam_total += rap.tam_total;
                                if (rap.es_agora)
                                    o_rap.es_agora = true;
                            }
                        }
                    }
                    db.CloseConnection();

                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();
                    #endregion 

                    //Muevo hoja cuatro después de la hoja uno
                    excelPackage.Workbook.Worksheets.MoveAfter(excelPackage.Workbook.Worksheets[3].Name,excelPackage.Workbook.Worksheets[0].Name);
                    //Muevo hoja cinco después de la hoja tres
                    excelPackage.Workbook.Worksheets.MoveAfter(excelPackage.Workbook.Worksheets[4].Name, excelPackage.Workbook.Worksheets[2].Name);

                    #region Resumen Agrupadas ES
                    //Añadimos nueva pestaña Resumen Agrupadas ES

                    workSheet = excelPackage.Workbook.Worksheets.Add("Resumen Agrupadas ES");
                    headerCells = workSheet.Cells[1, 1, 1, 9];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;

                    
                    //Cabecera
                    //Ponemos columnas directamente
                    workSheet.Cells[1, 1].Value = "NIF";
                    workSheet.Cells[1, 2].Value = "NOMBRE";
                    workSheet.Cells[1, 3].Value = "DIA AGRUPACION";
                    workSheet.Cells[1, 4].Value = "AGORA";
                    workSheet.Cells[1, 5].Value = "PENDIENTE MEDIDA";
                    workSheet.Cells[1, 6].Value = "OC CALCULABLE";
                    workSheet.Cells[1, 7].Value = "DC GENERADO SIN DI";
                    workSheet.Cells[1, 8].Value = "DI APARTADO";
                    workSheet.Cells[1, 9].Value = "TAM";

                    workSheet.View.FreezePanes(2, 1);
                    
                    //Recorremos el diccionario dic_resumen_agrupadas_ES y pintamos en el excel los registros
                    f = 1;
                    foreach (KeyValuePair<String, EndesaEntity.facturacion.ResumenAgrupadaPendiente> p in dic_resumen_agrupadas_ES)
                    {
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = p.Value.cif;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.razon_social;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.num_dia_agrupacion;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.es_agora?"SI":"NO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.pendiente_medida;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.oc_calculable;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.dc_generado_sin_di;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.di_apartado;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.tam_total;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        if (p.Value.tam_total >= Convert.ToDouble(param.GetValue("tam_resaltar_celda_agrupadas")))
                        {
                            workSheet.Cells[f, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(colorAviso);
                        }
                                
                        c++;
                    }

                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    allCells = workSheet.Cells[2, 1, f, c];
                    int[] columnas = { 2, 8 };
                    bool[] orden = { false, true };
                    allCells.Sort(columnas, orden);
                    
                   
                    #endregion

                    #region Resumen Agrupadas MT-BTE
                    //Añadimos nueva pestaña Resumen Agrupadas MT-BTE

                    workSheet = excelPackage.Workbook.Worksheets.Add("Resumen Agrupadas MT-BTE");
                    headerCells = workSheet.Cells[1, 1, 1, 9];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;


                    //Cabecera
                    //Ponemos columnas directamente
                    workSheet.Cells[1, 1].Value = "NIF";
                    workSheet.Cells[1, 2].Value = "NOMBRE";
                    workSheet.Cells[1, 3].Value = "DIA AGRUPACION";
                    workSheet.Cells[1, 4].Value = "AGORA";
                    workSheet.Cells[1, 5].Value = "PENDIENTE MEDIDA";
                    workSheet.Cells[1, 6].Value = "OC CALCULABLE";
                    workSheet.Cells[1, 7].Value = "DC GENERADO SIN DI";
                    workSheet.Cells[1, 8].Value = "DI APARTADO";
                    workSheet.Cells[1, 9].Value = "TAM";

                    workSheet.View.FreezePanes(2, 1);

                    //Recorremos el diccionario dic_resumen_agrupadas_MTBTE y pintamos en el excel los registros
                    f = 1;
                    foreach (KeyValuePair<String, EndesaEntity.facturacion.ResumenAgrupadaPendiente> p in dic_resumen_agrupadas_MTBTE)
                    {
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = p.Value.cif;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.razon_social;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.num_dia_agrupacion;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.es_agora ? "SI" : "NO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.pendiente_medida;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.oc_calculable;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.dc_generado_sin_di;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.di_apartado;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.tam_total;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        if (p.Value.tam_total >= Convert.ToDouble(param.GetValue("tam_resaltar_celda_agrupadas")))
                        {
                            workSheet.Cells[f, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(colorAviso);
                        }
                        c++;
                    }

                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    allCells = workSheet.Cells[2, 1, f, c];
                    allCells.Sort(columnas, orden);

                    #endregion

                    #region Resumen Agrupadas BTN
                    //Añadimos nueva pestaña Resumen Agrupadas BTN

                    workSheet = excelPackage.Workbook.Worksheets.Add("Resumen Agrupadas BTN");
                    headerCells = workSheet.Cells[1, 1, 1, 9];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;


                    //Cabecera
                    //Ponemos columnas directamente
                    workSheet.Cells[1, 1].Value = "NIF";
                    workSheet.Cells[1, 2].Value = "NOMBRE";
                    workSheet.Cells[1, 3].Value = "DIA AGRUPACION";
                    workSheet.Cells[1, 4].Value = "AGORA";
                    workSheet.Cells[1, 5].Value = "PENDIENTE MEDIDA";
                    workSheet.Cells[1, 6].Value = "OC CALCULABLE";
                    workSheet.Cells[1, 7].Value = "DC GENERADO SIN DI";
                    workSheet.Cells[1, 8].Value = "DI APARTADO";
                    workSheet.Cells[1, 9].Value = "TAM";

                    workSheet.View.FreezePanes(2, 1);

                    //Recorremos el diccionario dic_resumen_agrupadas_BTN y pintamos en el excel los registros
                    f = 1;
                    
                    foreach (KeyValuePair<String, EndesaEntity.facturacion.ResumenAgrupadaPendiente> p in dic_resumen_agrupadas_BTN)
                    {
                        f++;
                        c = 1;
                        workSheet.Cells[f, c].Value = p.Value.cif;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.razon_social;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.num_dia_agrupacion;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.es_agora ? "SI" : "NO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.pendiente_medida;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.oc_calculable;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.dc_generado_sin_di;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.di_apartado;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;
                        workSheet.Cells[f, c].Value = p.Value.tam_total;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        if (p.Value.tam_total >= Convert.ToDouble(param.GetValue("tam_resaltar_celda_agrupadas")))
                        {
                            workSheet.Cells[f, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(colorAviso);
                        }
                        c++;
                    }

                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();

                    allCells = workSheet.Cells[2, 1, f, c];
                    allCells.Sort(columnas, orden);

                    #endregion

                    //Guardarmos el informe pendiente SAP
                    excelPackage.SaveAs(fileInfo);


                    #region INFORME ESTADOS DESCARTADOS 03.C y 03.I
                    //Creamos otro fichero para informe estados descartados del pendiente SAP (a 19/11/2024: 03.C y 03.I)
                    FileInfo fileInfoEstadosDescartados = new FileInfo(ruta_salida_archivo_estados_descartados);
                    //ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                    ExcelPackage excelPackage_estados_descartados = new ExcelPackage(fileInfoEstadosDescartados);

                    #region Subestados descartados 03.C

                    f = 1;
                    c = 1;

                    workSheet = excelPackage_estados_descartados.Workbook.Worksheets.Add("Subestados 03.C");
                    headerCells = workSheet.Cells[1, 1, 1, 17];
                    headerFont = headerCells.Style.Font;

                    headerCells = workSheet.Cells[1, 1, 1, 30];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;


                    workSheet.View.FreezePanes(2, 1);

                    #region ESPAÑA 03.C

                    #region sentencia SQL ESPAÑA 03.C
                    strSql = "SELECT ps.cd_empr as EMPRESA, p.cd_segmento_ptg as SEGMENTO,  ps.de_tp_cli as TIPO_CLIENTE, ps.cd_nif_cif_cli as NIF, ps.tx_apell_cli as CLIENTE,"
                        + " ps.fh_alta_crto as FALTACONT, ps.fh_inicio_vers_crto as FPSERCON, p.cd_cups as cups20, p.id_instalacion as N_INSTALACION, ps.cd_tarifa_c as TARIFA,"
                        + " ps.cd_crto_comercial as CONTRATO, p.fh_periodo as MES, ps.de_empr_distdora_nombre as DISTRIBUIDORA, "
                        + " concat(p.cd_estado,' ',de.de_estado) as ESTADO, concat(p.cd_subestado,' ',ds.de_subestado) as SUBESTADO, '' as DIAS_ESTADO , p.TAM as TAM, IF(s.CUPS20 IS NULL, p.agora, 'S') AS agora,"
                        + " a.cd_tp_alarma AS TIPO_ALARMA, a.dcomenta AS ALARMA, AGR.nm_dia_agrup as DIA_AGRUPACION"
                        + " FROM ("

                        + " SELECT substr(cd_cups,1,20) as cd_cups, id_instalacion, cl_stro, id_crto_ext, cl_crto_ext,"
                        + " cd_empr_distdora, fh_desde, fh_hasta, fh_periodo, cd_estado, cd_subestado,"
                        + " lg_multimedida, cd_empr_titular, cd_ritmo_fact, cd_segmento_ptg, fh_envio,"
                        + " fec_act, cod_carga, TAM, agora, now()"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar "
                        + " WHERE cd_subestado = '03.C'"
                        + " AND fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "'"
                        + " ORDER BY cd_cups, fh_periodo DESC, fh_hasta DESC"

                        + " ) as p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_d_alarmasconfact_pendiente_facturar a ON"
                        + " a.cd_cups_ext = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"
                        + " LEFT OUTER JOIN fact.cm_sofisticados s on"
                        + " s.CUPS20 = p.cd_cups"
                        + " LEFT OUTER JOIN "
                        + " (SELECT LEFT(cd_cups, 20) AS cups20, nm_dia_agrup FROM fact.SAP_AGRUPADAS_AGCUPS_New WHERE fh_carga = (select max(fh_carga) from fact.SAP_AGRUPADAS_AGCUPS_New) GROUP BY cd_cups) AS AGR"
                        + " ON AGR.cups20 = p.cd_cups"
                        + " where p.fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "' ";
                    
                    //[26-02-2025] GUS: Añadimos códigos de empresas desde lista_empresas_ES en vez de hacerlo directamente [+ " p.cd_empr_titular in ('ES21','ES22') GROUP BY cups20, MES";]

                    firstOnly = true;

                    foreach (string p in lista_empresas_ES)
                    {
                        if (firstOnly)
                        {
                            strSql += " and p.cd_empr_titular in ("
                                + "'" + p + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + p + "'";
                    }

                    if (!firstOnly) //Si la lista tenía al menos un elemento cerramos sentencia correctamente.
                    {
                        strSql += ")";
                    }

                    strSql += " GROUP BY cups20, MES";
                    #endregion

                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();

                    
                    //Ponemos columnas directamente
                    
                    workSheet.Cells[1, 1].Value = "EMPRESA";
                    workSheet.Cells[1, 2].Value = "SEGMENTO";
                    workSheet.Cells[1, 3].Value = "TIPO CLIENTE";
                    workSheet.Cells[1, 4].Value = "NIF";
                    workSheet.Cells[1, 5].Value = "CLIENTE";
                    workSheet.Cells[1, 6].Value = "FALTACONT";
                    workSheet.Cells[1, 7].Value = "FPSERCON";
                    workSheet.Cells[1, 8].Value = "CUPS20";
                    workSheet.Cells[1, 9].Value = "Nº INSTALACIÓN";
                    workSheet.Cells[1, 10].Value = "TARIFA";
                    workSheet.Cells[1, 11].Value = "CONTRATO";
                    workSheet.Cells[1, 12].Value = "MES";
                    workSheet.Cells[1, 13].Value = "DISTRIBUIDORA";
                    workSheet.Cells[1, 14].Value = "ESTADO";
                    workSheet.Cells[1, 15].Value = "SUBESTADO";
                    workSheet.Cells[1, 16].Value = "DIAS ESTADO";
                    workSheet.Cells[1, 17].Value = "TAM";
                    workSheet.Cells[1, 18].Value = "ÁGORA";                 
                    workSheet.Cells[1, 19].Value = "TIPO ALARMA";    
                    workSheet.Cells[1, 20].Value = "ALARMA";    
                    workSheet.Cells[1, 21].Value = "DIA AGRUPACION";      
                    
                    //Añadimos registros de España
                    while (r.Read())
                    {
                        f++;
                        c = 1;

                        #region BUSCAR DATOS EN T_ED_H_PS_HIST
                        //Si el NIF está vacío supuestamente es porque el contrato está de baja y sale de la tabla cont.t_ed_h_ps
                        //Intentamos recuperar los campos EMPRESA, TIPO_CLIENTE, NIF, CLIENTE, TARIFA, CONTRATO Y DISTRIBUIDORA del histórico, cont.t_ed_h_ps_hist
                        Dictionary<string, string> datos_histo = new Dictionary<string, string>();
                        string dat = "";
                        if (r["NIF"] == System.DBNull.Value && r["CUPS20"] != System.DBNull.Value)
                        {
                            strSql_histo = "SELECT ps.cd_empr AS EMPRESA, ps.de_tp_cli AS TIPO_CLIENTE, ps.cd_nif_cif_cli AS NIF, ps.tx_apell_cli AS CLIENTE, ps.cd_tarifa_c AS TARIFA, ps.cd_crto_comercial AS CONTRATO, ps.de_empr_distdora_nombre AS DISTRIBUIDORA "
                                + "FROM cont.t_ed_h_ps_hist ps WHERE ps.cups20 = '" + r["CUPS20"] + "' ORDER BY ps.fh_act_dmco DESC LIMIT 1;";

                            db_histo = new MySQLDB(MySQLDB.Esquemas.GBL);
                            command_histo = new MySqlCommand(strSql_histo, db_histo.con);
                            r_histo = command_histo.ExecuteReader();
                            if (r_histo.Read())
                            {
                                if (r_histo["EMPRESA"] != System.DBNull.Value)
                                    datos_histo.Add("EMPRESA", r_histo["EMPRESA"].ToString());
                                if (r_histo["TIPO_CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("TIPO_CLIENTE", r_histo["TIPO_CLIENTE"].ToString());
                                if (r_histo["NIF"] != System.DBNull.Value)
                                    datos_histo.Add("NIF", r_histo["NIF"].ToString());
                                if (r_histo["CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("CLIENTE", r_histo["CLIENTE"].ToString());
                                if (r_histo["TARIFA"] != System.DBNull.Value)
                                    datos_histo.Add("TARIFA", r_histo["TARIFA"].ToString());
                                if (r_histo["CONTRATO"] != System.DBNull.Value)
                                    datos_histo.Add("CONTRATO", r_histo["CONTRATO"].ToString());
                                if (r_histo["DISTRIBUIDORA"] != System.DBNull.Value)
                                    datos_histo.Add("DISTRIBUIDORA", r_histo["DISTRIBUIDORA"].ToString());
                            }

                            db_histo.CloseConnection();
                        }
                        #endregion

                        //Creamos objeto ResumenAgrupadaPendiente para ir almacenando info y al final, si procede, añadir a dic dic_resumen_agrupadas_ES
                        //EndesaEntity.facturacion.ResumenAgrupadaPendiente rap = new EndesaEntity.facturacion.ResumenAgrupadaPendiente();
                        //tiene_num_dia_agrupacion = false;


                        if (r["EMPRESA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["EMPRESA"].ToString();
                        else if (datos_histo.TryGetValue("EMPRESA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["SEGMENTO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["SEGMENTO"].ToString();
                        c++;

                        if (r["TIPO_CLIENTE"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TIPO_CLIENTE"].ToString();
                        else if (datos_histo.TryGetValue("TIPO_CLIENTE", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["NIF"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["NIF"].ToString();
                            //rap.cif = r["NIF"].ToString();
                        }
                        else if (datos_histo.TryGetValue("NIF", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            //rap.cif = dat;
                        }
                        c++;

                        if (r["CLIENTE"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["CLIENTE"].ToString();
                            //rap.razon_social = r["CLIENTE"].ToString();
                        }
                        else if (datos_histo.TryGetValue("CLIENTE", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            //rap.razon_social = dat;
                        }
                        c++;

                        if (r["FALTACONT"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["FALTACONT"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["FPSERCON"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["FPSERCON"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["CUPS20"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CUPS20"].ToString();
                        c++;

                        if (r["N_INSTALACION"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["N_INSTALACION"].ToString();
                        c++;

                        if (r["TARIFA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                        else if (datos_histo.TryGetValue("TARIFA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["CONTRATO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CONTRATO"].ToString();
                        else if (datos_histo.TryGetValue("CONTRATO", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["MES"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["MES"]);
                            aniomes = Convert.ToInt32(r["MES"]);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                        }
                        c++;

                        if (r["DISTRIBUIDORA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["DISTRIBUIDORA"].ToString();
                        else if (datos_histo.TryGetValue("DISTRIBUIDORA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["ESTADO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["ESTADO"].ToString();
                        c++;

                        if (r["SUBESTADO"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["SUBESTADO"].ToString();
                            //switch (r["SUBESTADO"].ToString().Substring(0, 2))
                            //{
                            //    case "01":
                            //        rap.pendiente_medida = 1;
                            //        break;
                            //    case "02":
                            //        rap.oc_calculable = 1;
                            //        break;
                            //    case "03":
                            //        rap.dc_generado_sin_di = 1;
                            //        break;
                            //    case "04":
                            //        rap.di_apartado = 1;
                            //        break;

                            //}
                        }
                        c++;

                        //if (r["lg_multimedida"] != System.DBNull.Value)
                        //{
                        //    workSheet.Cells[f, c].Value = r["lg_multimedida"].ToString();

                        //}
                        //else
                        //{
                        //    workSheet.Cells[f, c].Value = "N";
                        //}

                        workSheet.Cells[f, c].Value = GetDiasCUPSSubEstado(r["CUPS20"].ToString(), r["SUBESTADO"].ToString().Split(' ')[0]);

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["TAM"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            //rap.tam_total = rap.tam_total + Convert.ToDouble(r["TAM"]);
                        }
                        c++;

                        // NO MOSTRAMOS EN ESTE CASO LOS MESES NI IMPORTES PENDIENTES
                        //if (r["MES"] != System.DBNull.Value)
                        //{
                        //    //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                        //    meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                        //    workSheet.Cells[f, c].Value = meses_pdtes;
                        //    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //}
                        //c++;

                        //if (r["MES"] != System.DBNull.Value && r["TAM"] != System.DBNull.Value)
                        //{
                        //    //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                        //    meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                        //    workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]) * meses_pdtes;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        //}
                        //else
                        //{
                        //    c++;
                        //}

                        if (r["AGORA"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["AGORA"].ToString();
                            //if (r["AGORA"].ToString() == "S")
                            //    rap.es_agora = true;
                        }

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["TIPO_ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TIPO_ALARMA"].ToString();

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["ALARMA"].ToString();

                        c++;

                        if (r["DIA_AGRUPACION"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["DIA_AGRUPACION"].ToString();
                            //rap.num_dia_agrupacion = Convert.ToInt16(r["DIA_AGRUPACION"]);
                            //tiene_num_dia_agrupacion = true;
                        }

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        
                    }
                    db.CloseConnection();


                    #endregion 


                    #region PORTUGAL 03.C

                    #region sentencia SQL PORTUGAL 03.C
                    strSql = "SELECT ps.cd_empr as EMPRESA, p.cd_segmento_ptg as SEGMENTO,  ps.de_tp_cli as TIPO_CLIENTE, ps.cd_nif_cif_cli as NIF, ps.tx_apell_cli as CLIENTE,"
                        + " ps.fh_alta_crto as FALTACONT, ps.fh_inicio_vers_crto as FPSERCON, p.cd_cups as cups20, p.id_instalacion as N_INSTALACION, ps.cd_tarifa_c as TARIFA,"
                        + " ps.cd_crto_comercial as CONTRATO, p.fh_periodo as MES, ps.de_empr_distdora_nombre as DISTRIBUIDORA, "
                        + " concat(p.cd_estado,' ',de.de_estado) as ESTADO, concat(p.cd_subestado,' ',ds.de_subestado) as SUBESTADO, '' as DIAS_ESTADO , p.TAM as TAM, IF(s.CUPS20 IS NULL, p.agora, 'S') AS agora,"
                        + " a.cd_tp_alarma AS TIPO_ALARMA, a.dcomenta AS ALARMA, AGR.nm_dia_agrup as DIA_AGRUPACION"
                        + " FROM ("

                        + " SELECT substr(cd_cups,1,20) as cd_cups, id_instalacion, cl_stro, id_crto_ext, cl_crto_ext,"
                        + " cd_empr_distdora, fh_desde, fh_hasta, fh_periodo, cd_estado, cd_subestado,"
                        + " lg_multimedida, cd_empr_titular, cd_ritmo_fact, cd_segmento_ptg, fh_envio,"
                        + " fec_act, cod_carga, TAM, agora, now()"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar "
                        + " WHERE cd_subestado = '03.C'"
                        + " AND fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "'"
                        + " ORDER BY cd_cups, fh_periodo DESC, fh_hasta DESC"

                        + " ) as p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps_pt ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_d_alarmasconfact_pendiente_facturar a ON"
                        + " a.cd_cups_ext = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"
                        + " LEFT OUTER JOIN fact.cm_sofisticados s on"
                        + " s.CUPS20 = p.cd_cups"
                        + " LEFT OUTER JOIN "
                        + " (SELECT LEFT(cd_cups, 20) AS cups20, nm_dia_agrup FROM fact.SAP_AGRUPADAS_AGCUPS_New WHERE fh_carga = (select max(fh_carga) from fact.SAP_AGRUPADAS_AGCUPS_New) GROUP BY cd_cups) AS AGR"
                        + " ON AGR.cups20 = p.cd_cups"
                        + " where p.fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "' ";
                      
                    //[26-02-2025] GUS: Añadimos códigos de empresas desde lista_empresas_ES
                    // y segmentos de Portugal desde lista_segmentos_MT_BTE + lista_segmentos_BTN
                    // en vez de hacerlo directamente [+ " (p.cd_empr_titular = 'PT1Q' and p.cd_segmento_ptg in ('MT','BTE','BTN')) GROUP BY cups20, MES";]

                    //Lista empresas Portugal
                    firstOnly = true;
                    foreach (string p in lista_empresas_PT)
                    {
                        if (firstOnly)
                        {
                            strSql += " and p.cd_empr_titular in ("
                                + "'" + p + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + p + "'";
                    }
                    if (!firstOnly) //Si la lista tenía al menos un elemento cerramos sentencia correctamente.
                    {
                        strSql += ")";
                    }

                    //Lista segmentos Portugal (dos listas)
                    firstOnly = true;
                    foreach (string p in lista_segmentos_MT_BTE)
                    {
                        if (firstOnly)
                        {
                            strSql += " and (p.cd_segmento_ptg in ("
                                + "'" + p + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + p + "'";
                    }
                    foreach (string p in lista_segmentos_BTN)
                    {
                        if (firstOnly)
                        {
                            strSql += " and (p.cd_segmento_ptg in ("
                                + "'" + p + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + p + "'";
                    }

                    if (!firstOnly) //Si la lista tenía al menos un elemento cerramos sentencia correctamente, incluimos registros sin segmento informado.
                    {
                        strSql += ") or p.cd_segmento_ptg is null)";
                    }
                    else
                        strSql += " and p.cd_segmento_ptg is null";

                    strSql += " GROUP BY cups20, MES";

                    #endregion

                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();

                    

                    //Ponemos columnas directamente

                    workSheet.Cells[1, 1].Value = "EMPRESA";
                    workSheet.Cells[1, 2].Value = "SEGMENTO";
                    workSheet.Cells[1, 3].Value = "TIPO CLIENTE";
                    workSheet.Cells[1, 4].Value = "NIF";
                    workSheet.Cells[1, 5].Value = "CLIENTE";
                    workSheet.Cells[1, 6].Value = "FALTACONT";
                    workSheet.Cells[1, 7].Value = "FPSERCON";
                    workSheet.Cells[1, 8].Value = "CUPS20";
                    workSheet.Cells[1, 9].Value = "Nº INSTALACIÓN";
                    workSheet.Cells[1, 10].Value = "TARIFA";
                    workSheet.Cells[1, 11].Value = "CONTRATO";
                    workSheet.Cells[1, 12].Value = "MES";
                    workSheet.Cells[1, 13].Value = "DISTRIBUIDORA";
                    workSheet.Cells[1, 14].Value = "ESTADO";
                    workSheet.Cells[1, 15].Value = "SUBESTADO";
                    workSheet.Cells[1, 16].Value = "DIAS ESTADO";
                    workSheet.Cells[1, 17].Value = "TAM";
                    workSheet.Cells[1, 18].Value = "ÁGORA";
                    workSheet.Cells[1, 19].Value = "TIPO ALARMA";
                    workSheet.Cells[1, 20].Value = "ALARMA";
                    workSheet.Cells[1, 21].Value = "DIA AGRUPACION";

                    //Añadimos registros de España
                    while (r.Read())
                    {
                        f++;
                        c = 1;

                        #region BUSCAR DATOS EN T_ED_H_PS_HIST
                        //Si el NIF está vacío supuestamente es porque el contrato está de baja y sale de la tabla cont.t_ed_h_ps
                        //Intentamos recuperar los campos EMPRESA, TIPO_CLIENTE, NIF, CLIENTE, TARIFA, CONTRATO Y DISTRIBUIDORA del histórico, cont.t_ed_h_ps_hist
                        Dictionary<string, string> datos_histo = new Dictionary<string, string>();
                        string dat = "";
                        if (r["NIF"] == System.DBNull.Value && r["CUPS20"] != System.DBNull.Value)
                        {
                            strSql_histo = "SELECT ps.cd_empr AS EMPRESA, ps.de_tp_cli AS TIPO_CLIENTE, ps.cd_nif_cif_cli AS NIF, ps.tx_apell_cli AS CLIENTE, ps.cd_tarifa_c AS TARIFA, ps.cd_crto_comercial AS CONTRATO, ps.de_empr_distdora_nombre AS DISTRIBUIDORA "
                                + "FROM cont.t_ed_h_ps_pt_hist ps WHERE ps.cups20 = '" + r["CUPS20"] + "' ORDER BY ps.fh_act_dmco DESC LIMIT 1;";

                            db_histo = new MySQLDB(MySQLDB.Esquemas.GBL);
                            command_histo = new MySqlCommand(strSql_histo, db_histo.con);
                            r_histo = command_histo.ExecuteReader();
                            if (r_histo.Read())
                            {
                                if (r_histo["EMPRESA"] != System.DBNull.Value)
                                    datos_histo.Add("EMPRESA", r_histo["EMPRESA"].ToString());
                                if (r_histo["TIPO_CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("TIPO_CLIENTE", r_histo["TIPO_CLIENTE"].ToString());
                                if (r_histo["NIF"] != System.DBNull.Value)
                                    datos_histo.Add("NIF", r_histo["NIF"].ToString());
                                if (r_histo["CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("CLIENTE", r_histo["CLIENTE"].ToString());
                                if (r_histo["TARIFA"] != System.DBNull.Value)
                                    datos_histo.Add("TARIFA", r_histo["TARIFA"].ToString());
                                if (r_histo["CONTRATO"] != System.DBNull.Value)
                                    datos_histo.Add("CONTRATO", r_histo["CONTRATO"].ToString());
                                if (r_histo["DISTRIBUIDORA"] != System.DBNull.Value)
                                    datos_histo.Add("DISTRIBUIDORA", r_histo["DISTRIBUIDORA"].ToString());
                            }

                            db_histo.CloseConnection();
                        }
                        #endregion

                        //Creamos objeto ResumenAgrupadaPendiente para ir almacenando info y al final, si procede, añadir a dic dic_resumen_agrupadas_ES
                        //EndesaEntity.facturacion.ResumenAgrupadaPendiente rap = new EndesaEntity.facturacion.ResumenAgrupadaPendiente();
                        //tiene_num_dia_agrupacion = false;


                        if (r["EMPRESA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["EMPRESA"].ToString();
                        else if (datos_histo.TryGetValue("EMPRESA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["SEGMENTO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["SEGMENTO"].ToString();
                        else
                            workSheet.Cells[f, c].Value = "NULO";
                        c++;

                        if (r["TIPO_CLIENTE"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TIPO_CLIENTE"].ToString();
                        else if (datos_histo.TryGetValue("TIPO_CLIENTE", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["NIF"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["NIF"].ToString();
                            //rap.cif = r["NIF"].ToString();
                        }
                        else if (datos_histo.TryGetValue("NIF", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            //rap.cif = dat;
                        }
                        c++;

                        if (r["CLIENTE"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["CLIENTE"].ToString();
                            //rap.razon_social = r["CLIENTE"].ToString();
                        }
                        else if (datos_histo.TryGetValue("CLIENTE", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            //rap.razon_social = dat;
                        }
                        c++;

                        if (r["FALTACONT"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["FALTACONT"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["FPSERCON"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["FPSERCON"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["CUPS20"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CUPS20"].ToString();
                        c++;

                        if (r["N_INSTALACION"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["N_INSTALACION"].ToString();
                        c++;

                        if (r["TARIFA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                        else if (datos_histo.TryGetValue("TARIFA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["CONTRATO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CONTRATO"].ToString();
                        else if (datos_histo.TryGetValue("CONTRATO", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["MES"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["MES"]);
                            aniomes = Convert.ToInt32(r["MES"]);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                        }
                        c++;

                        if (r["DISTRIBUIDORA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["DISTRIBUIDORA"].ToString();
                        else if (datos_histo.TryGetValue("DISTRIBUIDORA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["ESTADO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["ESTADO"].ToString();
                        c++;

                        if (r["SUBESTADO"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["SUBESTADO"].ToString();
                            //switch (r["SUBESTADO"].ToString().Substring(0, 2))
                            //{
                            //    case "01":
                            //        rap.pendiente_medida = 1;
                            //        break;
                            //    case "02":
                            //        rap.oc_calculable = 1;
                            //        break;
                            //    case "03":
                            //        rap.dc_generado_sin_di = 1;
                            //        break;
                            //    case "04":
                            //        rap.di_apartado = 1;
                            //        break;

                            //}
                        }
                        c++;

                        //if (r["lg_multimedida"] != System.DBNull.Value)
                        //{
                        //    workSheet.Cells[f, c].Value = r["lg_multimedida"].ToString();

                        //}
                        //else
                        //{
                        //    workSheet.Cells[f, c].Value = "N";
                        //}

                        workSheet.Cells[f, c].Value = GetDiasCUPSSubEstado(r["CUPS20"].ToString(), r["SUBESTADO"].ToString().Split(' ')[0]);

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["TAM"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            //rap.tam_total = rap.tam_total + Convert.ToDouble(r["TAM"]);
                        }
                        c++;

                        // NO MOSTRAMOS EN ESTE CASO LOS MESES NI IMPORTES PENDIENTES
                        //if (r["MES"] != System.DBNull.Value)
                        //{
                        //    //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                        //    meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                        //    workSheet.Cells[f, c].Value = meses_pdtes;
                        //    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //}
                        //c++;

                        //if (r["MES"] != System.DBNull.Value && r["TAM"] != System.DBNull.Value)
                        //{
                        //    //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                        //    meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                        //    workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]) * meses_pdtes;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        //}
                        //else
                        //{
                        //    c++;
                        //}

                        if (r["AGORA"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["AGORA"].ToString();
                            //if (r["AGORA"].ToString() == "S")
                            //    rap.es_agora = true;
                        }

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["TIPO_ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TIPO_ALARMA"].ToString();

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["ALARMA"].ToString();

                        c++;

                        if (r["DIA_AGRUPACION"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["DIA_AGRUPACION"].ToString();
                            //rap.num_dia_agrupacion = Convert.ToInt16(r["DIA_AGRUPACION"]);
                            //tiene_num_dia_agrupacion = true;
                        }

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        
                    }
                    db.CloseConnection();


                    #endregion


                    #endregion


                    headerCells = workSheet.Cells[1, 1, 1, c];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;
                    allCells = workSheet.Cells[1, 1, f, c];

                    allCells.AutoFitColumns();
                    workSheet.Cells[1,1,1,c].AutoFilter = true;




                    #region Subestados descartados 03.I

                    f = 1;
                    c = 1;

                    workSheet = excelPackage_estados_descartados.Workbook.Worksheets.Add("Subestados 03.I");
                    headerCells = workSheet.Cells[1, 1, 1, 17];
                    headerFont = headerCells.Style.Font;

                    headerCells = workSheet.Cells[1, 1, 1, 30];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;


                    workSheet.View.FreezePanes(2, 1);

                    #region ESPAÑA 03.I

                    #region sentencia SQL ESPAÑA 03.I
                    strSql = "SELECT ps.cd_empr as EMPRESA, p.cd_segmento_ptg as SEGMENTO,  ps.de_tp_cli as TIPO_CLIENTE, ps.cd_nif_cif_cli as NIF, ps.tx_apell_cli as CLIENTE,"
                        + " ps.fh_alta_crto as FALTACONT, ps.fh_inicio_vers_crto as FPSERCON, p.cd_cups as cups20, p.id_instalacion as N_INSTALACION, ps.cd_tarifa_c as TARIFA,"
                        + " ps.cd_crto_comercial as CONTRATO, p.fh_periodo as MES, ps.de_empr_distdora_nombre as DISTRIBUIDORA, "
                        + " concat(p.cd_estado,' ',de.de_estado) as ESTADO, concat(p.cd_subestado,' ',ds.de_subestado) as SUBESTADO, '' as DIAS_ESTADO , p.TAM as TAM, IF(s.CUPS20 IS NULL, p.agora, 'S') AS agora,"
                        + " a.cd_tp_alarma AS TIPO_ALARMA, a.dcomenta AS ALARMA, AGR.nm_dia_agrup as DIA_AGRUPACION"
                        + " FROM ("

                        + " SELECT substr(cd_cups,1,20) as cd_cups, id_instalacion, cl_stro, id_crto_ext, cl_crto_ext,"
                        + " cd_empr_distdora, fh_desde, fh_hasta, fh_periodo, cd_estado, cd_subestado,"
                        + " lg_multimedida, cd_empr_titular, cd_ritmo_fact, cd_segmento_ptg, fh_envio,"
                        + " fec_act, cod_carga, TAM, agora, now()"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar "
                        + " WHERE cd_subestado = '03.I'"
                        + " AND fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "'"
                        + " ORDER BY cd_cups, fh_periodo DESC, fh_hasta DESC"

                        + " ) as p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_d_alarmasconfact_pendiente_facturar a ON"
                        + " a.cd_cups_ext = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"
                        + " LEFT OUTER JOIN fact.cm_sofisticados s on"
                        + " s.CUPS20 = p.cd_cups"
                        + " LEFT OUTER JOIN "
                        + " (SELECT LEFT(cd_cups, 20) AS cups20, nm_dia_agrup FROM fact.SAP_AGRUPADAS_AGCUPS_New WHERE fh_carga = (select max(fh_carga) from fact.SAP_AGRUPADAS_AGCUPS_New) GROUP BY cd_cups) AS AGR"
                        + " ON AGR.cups20 = p.cd_cups"
                        + " where p.fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "' ";
                        
                        
                    //[26-02-2025] GUS: Añadimos códigos de empresas desde lista_empresas_ES en vez de hacerlo directamente [+ " p.cd_empr_titular in ('ES21','ES22') GROUP BY cups20, MES";]

                    firstOnly = true;

                    foreach (string p in lista_empresas_ES)
                    {
                        if (firstOnly)
                        {
                            strSql += " and p.cd_empr_titular in ("
                                + "'" + p + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + p + "'";
                    }

                    if (!firstOnly) //Si la lista tenía al menos un elemento cerramos sentencia correctamente.
                    {
                        strSql += ")";
                    }

                    strSql += " GROUP BY cups20, MES";
                    #endregion

                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();


                    //Ponemos columnas directamente

                    workSheet.Cells[1, 1].Value = "EMPRESA";
                    workSheet.Cells[1, 2].Value = "SEGMENTO";
                    workSheet.Cells[1, 3].Value = "TIPO CLIENTE";
                    workSheet.Cells[1, 4].Value = "NIF";
                    workSheet.Cells[1, 5].Value = "CLIENTE";
                    workSheet.Cells[1, 6].Value = "FALTACONT";
                    workSheet.Cells[1, 7].Value = "FPSERCON";
                    workSheet.Cells[1, 8].Value = "CUPS20";
                    workSheet.Cells[1, 9].Value = "Nº INSTALACIÓN";
                    workSheet.Cells[1, 10].Value = "TARIFA";
                    workSheet.Cells[1, 11].Value = "CONTRATO";
                    workSheet.Cells[1, 12].Value = "MES";
                    workSheet.Cells[1, 13].Value = "DISTRIBUIDORA";
                    workSheet.Cells[1, 14].Value = "ESTADO";
                    workSheet.Cells[1, 15].Value = "SUBESTADO";
                    workSheet.Cells[1, 16].Value = "DIAS ESTADO";
                    workSheet.Cells[1, 17].Value = "TAM";
                    workSheet.Cells[1, 18].Value = "ÁGORA";
                    workSheet.Cells[1, 19].Value = "TIPO ALARMA";
                    workSheet.Cells[1, 20].Value = "ALARMA";
                    workSheet.Cells[1, 21].Value = "DIA AGRUPACION";

                    //Añadimos registros de España
                    while (r.Read())
                    {
                        f++;
                        c = 1;

                        #region BUSCAR DATOS EN T_ED_H_PS_HIST
                        //Si el NIF está vacío supuestamente es porque el contrato está de baja y sale de la tabla cont.t_ed_h_ps
                        //Intentamos recuperar los campos EMPRESA, TIPO_CLIENTE, NIF, CLIENTE, TARIFA, CONTRATO Y DISTRIBUIDORA del histórico, cont.t_ed_h_ps_hist
                        Dictionary<string, string> datos_histo = new Dictionary<string, string>();
                        string dat = "";
                        if (r["NIF"] == System.DBNull.Value && r["CUPS20"] != System.DBNull.Value)
                        {
                            strSql_histo = "SELECT ps.cd_empr AS EMPRESA, ps.de_tp_cli AS TIPO_CLIENTE, ps.cd_nif_cif_cli AS NIF, ps.tx_apell_cli AS CLIENTE, ps.cd_tarifa_c AS TARIFA, ps.cd_crto_comercial AS CONTRATO, ps.de_empr_distdora_nombre AS DISTRIBUIDORA "
                                + "FROM cont.t_ed_h_ps_hist ps WHERE ps.cups20 = '" + r["CUPS20"] + "' ORDER BY ps.fh_act_dmco DESC LIMIT 1;";

                            db_histo = new MySQLDB(MySQLDB.Esquemas.GBL);
                            command_histo = new MySqlCommand(strSql_histo, db_histo.con);
                            r_histo = command_histo.ExecuteReader();
                            if (r_histo.Read())
                            {
                                if (r_histo["EMPRESA"] != System.DBNull.Value)
                                    datos_histo.Add("EMPRESA", r_histo["EMPRESA"].ToString());
                                if (r_histo["TIPO_CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("TIPO_CLIENTE", r_histo["TIPO_CLIENTE"].ToString());
                                if (r_histo["NIF"] != System.DBNull.Value)
                                    datos_histo.Add("NIF", r_histo["NIF"].ToString());
                                if (r_histo["CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("CLIENTE", r_histo["CLIENTE"].ToString());
                                if (r_histo["TARIFA"] != System.DBNull.Value)
                                    datos_histo.Add("TARIFA", r_histo["TARIFA"].ToString());
                                if (r_histo["CONTRATO"] != System.DBNull.Value)
                                    datos_histo.Add("CONTRATO", r_histo["CONTRATO"].ToString());
                                if (r_histo["DISTRIBUIDORA"] != System.DBNull.Value)
                                    datos_histo.Add("DISTRIBUIDORA", r_histo["DISTRIBUIDORA"].ToString());
                            }

                            db_histo.CloseConnection();
                        }
                        #endregion

                        //Creamos objeto ResumenAgrupadaPendiente para ir almacenando info y al final, si procede, añadir a dic dic_resumen_agrupadas_ES
                        //EndesaEntity.facturacion.ResumenAgrupadaPendiente rap = new EndesaEntity.facturacion.ResumenAgrupadaPendiente();
                        //tiene_num_dia_agrupacion = false;


                        if (r["EMPRESA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["EMPRESA"].ToString();
                        else if (datos_histo.TryGetValue("EMPRESA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["SEGMENTO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["SEGMENTO"].ToString();
                        c++;

                        if (r["TIPO_CLIENTE"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TIPO_CLIENTE"].ToString();
                        else if (datos_histo.TryGetValue("TIPO_CLIENTE", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["NIF"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["NIF"].ToString();
                            //rap.cif = r["NIF"].ToString();
                        }
                        else if (datos_histo.TryGetValue("NIF", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            //rap.cif = dat;
                        }
                        c++;

                        if (r["CLIENTE"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["CLIENTE"].ToString();
                            //rap.razon_social = r["CLIENTE"].ToString();
                        }
                        else if (datos_histo.TryGetValue("CLIENTE", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            //rap.razon_social = dat;
                        }
                        c++;

                        if (r["FALTACONT"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["FALTACONT"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["FPSERCON"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["FPSERCON"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["CUPS20"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CUPS20"].ToString();
                        c++;

                        if (r["N_INSTALACION"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["N_INSTALACION"].ToString();
                        c++;

                        if (r["TARIFA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                        else if (datos_histo.TryGetValue("TARIFA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["CONTRATO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CONTRATO"].ToString();
                        else if (datos_histo.TryGetValue("CONTRATO", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["MES"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["MES"]);
                            aniomes = Convert.ToInt32(r["MES"]);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                        }
                        c++;

                        if (r["DISTRIBUIDORA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["DISTRIBUIDORA"].ToString();
                        else if (datos_histo.TryGetValue("DISTRIBUIDORA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["ESTADO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["ESTADO"].ToString();
                        c++;

                        if (r["SUBESTADO"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["SUBESTADO"].ToString();
                            //switch (r["SUBESTADO"].ToString().Substring(0, 2))
                            //{
                            //    case "01":
                            //        rap.pendiente_medida = 1;
                            //        break;
                            //    case "02":
                            //        rap.oc_calculable = 1;
                            //        break;
                            //    case "03":
                            //        rap.dc_generado_sin_di = 1;
                            //        break;
                            //    case "04":
                            //        rap.di_apartado = 1;
                            //        break;

                            //}
                        }
                        c++;

                        //if (r["lg_multimedida"] != System.DBNull.Value)
                        //{
                        //    workSheet.Cells[f, c].Value = r["lg_multimedida"].ToString();

                        //}
                        //else
                        //{
                        //    workSheet.Cells[f, c].Value = "N";
                        //}

                        workSheet.Cells[f, c].Value = GetDiasCUPSSubEstado(r["CUPS20"].ToString(), r["SUBESTADO"].ToString().Split(' ')[0]);

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["TAM"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            //rap.tam_total = rap.tam_total + Convert.ToDouble(r["TAM"]);
                        }
                        c++;

                        // NO MOSTRAMOS EN ESTE CASO LOS MESES NI IMPORTES PENDIENTES
                        //if (r["MES"] != System.DBNull.Value)
                        //{
                        //    //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                        //    meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                        //    workSheet.Cells[f, c].Value = meses_pdtes;
                        //    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //}
                        //c++;

                        //if (r["MES"] != System.DBNull.Value && r["TAM"] != System.DBNull.Value)
                        //{
                        //    //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                        //    meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                        //    workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]) * meses_pdtes;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        //}
                        //else
                        //{
                        //    c++;
                        //}

                        if (r["AGORA"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["AGORA"].ToString();
                            //if (r["AGORA"].ToString() == "S")
                            //    rap.es_agora = true;
                        }

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["TIPO_ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TIPO_ALARMA"].ToString();

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["ALARMA"].ToString();

                        c++;

                        if (r["DIA_AGRUPACION"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["DIA_AGRUPACION"].ToString();
                            //rap.num_dia_agrupacion = Convert.ToInt16(r["DIA_AGRUPACION"]);
                            //tiene_num_dia_agrupacion = true;
                        }

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        
                    }
                    db.CloseConnection();


                    #endregion


                    #region PORTUGAL 03.I

                    #region sentencia SQL PORTUGAL 03.I
                    strSql = "SELECT ps.cd_empr as EMPRESA, p.cd_segmento_ptg as SEGMENTO,  ps.de_tp_cli as TIPO_CLIENTE, ps.cd_nif_cif_cli as NIF, ps.tx_apell_cli as CLIENTE,"
                        + " ps.fh_alta_crto as FALTACONT, ps.fh_inicio_vers_crto as FPSERCON, p.cd_cups as cups20, p.id_instalacion as N_INSTALACION, ps.cd_tarifa_c as TARIFA,"
                        + " ps.cd_crto_comercial as CONTRATO, p.fh_periodo as MES, ps.de_empr_distdora_nombre as DISTRIBUIDORA, "
                        + " concat(p.cd_estado,' ',de.de_estado) as ESTADO, concat(p.cd_subestado,' ',ds.de_subestado) as SUBESTADO, '' as DIAS_ESTADO , p.TAM as TAM, IF(s.CUPS20 IS NULL, p.agora, 'S') AS agora,"
                        + " a.cd_tp_alarma AS TIPO_ALARMA, a.dcomenta AS ALARMA, AGR.nm_dia_agrup as DIA_AGRUPACION"
                        + " FROM ("

                        + " SELECT substr(cd_cups,1,20) as cd_cups, id_instalacion, cl_stro, id_crto_ext, cl_crto_ext,"
                        + " cd_empr_distdora, fh_desde, fh_hasta, fh_periodo, cd_estado, cd_subestado,"
                        + " lg_multimedida, cd_empr_titular, cd_ritmo_fact, cd_segmento_ptg, fh_envio,"
                        + " fec_act, cod_carga, TAM, agora, now()"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar "
                        + " WHERE cd_subestado = '03.I'"
                        + " AND fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "'"
                        + " ORDER BY cd_cups, fh_periodo DESC, fh_hasta DESC"

                        + " ) as p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps_pt ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_d_alarmasconfact_pendiente_facturar a ON"
                        + " a.cd_cups_ext = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"
                        + " LEFT OUTER JOIN fact.cm_sofisticados s on"
                        + " s.CUPS20 = p.cd_cups"
                        + " LEFT OUTER JOIN "
                        + " (SELECT LEFT(cd_cups, 20) AS cups20, nm_dia_agrup FROM fact.SAP_AGRUPADAS_AGCUPS_New WHERE fh_carga = (select max(fh_carga) from fact.SAP_AGRUPADAS_AGCUPS_New) GROUP BY cd_cups) AS AGR"
                        + " ON AGR.cups20 = p.cd_cups"
                        + " where p.fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "' ";
                        //+ " (p.cd_empr_titular = 'PT1Q' and p.cd_segmento_ptg in ('MT','BTE','BTN')) GROUP BY cups20, MES";

                    //[26-02-2025] GUS: Añadimos códigos de empresas desde lista_empresas_ES
                    // y segmentos de Portugal desde lista_segmentos_MT_BTE + lista_segmentos_BTN
                    // en vez de hacerlo directamente [+ " (p.cd_empr_titular = 'PT1Q' and p.cd_segmento_ptg in ('MT','BTE','BTN')) GROUP BY cups20, MES";]

                    //Lista de empresas Portugal
                    firstOnly = true;
                    foreach (string p in lista_empresas_PT)
                    {
                        if (firstOnly)
                        {
                            strSql += " and p.cd_empr_titular in ("
                                + "'" + p + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + p + "'";
                    }
                    if (!firstOnly) //Si la lista tenía al menos un elemento cerramos sentencia correctamente.
                    {
                        strSql += ")";
                    }

                    //Lista de segmentos Portugal (las dos listas)
                    firstOnly = true;
                    foreach (string p in lista_segmentos_MT_BTE)
                    {
                        if (firstOnly)
                        {
                            strSql += " and (p.cd_segmento_ptg in ("
                                + "'" + p + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + p + "'";
                    }
                    foreach (string p in lista_segmentos_BTN)
                    {
                        if (firstOnly)
                        {
                            strSql += " and (p.cd_segmento_ptg in ("
                                + "'" + p + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + p + "'";
                    }

                    if (!firstOnly) //Si la lista tenía al menos un elemento cerramos sentencia correctamente, incluimos registros sin segmento informado.
                    {
                        strSql += ") or p.cd_segmento_ptg is null)";
                    }
                    else
                        strSql += " and p.cd_segmento_ptg is null";

                    strSql += " GROUP BY cups20, MES";
                    #endregion

                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();



                    //Ponemos columnas directamente

                    workSheet.Cells[1, 1].Value = "EMPRESA";
                    workSheet.Cells[1, 2].Value = "SEGMENTO";
                    workSheet.Cells[1, 3].Value = "TIPO CLIENTE";
                    workSheet.Cells[1, 4].Value = "NIF";
                    workSheet.Cells[1, 5].Value = "CLIENTE";
                    workSheet.Cells[1, 6].Value = "FALTACONT";
                    workSheet.Cells[1, 7].Value = "FPSERCON";
                    workSheet.Cells[1, 8].Value = "CUPS20";
                    workSheet.Cells[1, 9].Value = "Nº INSTALACIÓN";
                    workSheet.Cells[1, 10].Value = "TARIFA";
                    workSheet.Cells[1, 11].Value = "CONTRATO";
                    workSheet.Cells[1, 12].Value = "MES";
                    workSheet.Cells[1, 13].Value = "DISTRIBUIDORA";
                    workSheet.Cells[1, 14].Value = "ESTADO";
                    workSheet.Cells[1, 15].Value = "SUBESTADO";
                    workSheet.Cells[1, 16].Value = "DIAS ESTADO";
                    workSheet.Cells[1, 17].Value = "TAM";
                    workSheet.Cells[1, 18].Value = "ÁGORA";
                    workSheet.Cells[1, 19].Value = "TIPO ALARMA";
                    workSheet.Cells[1, 20].Value = "ALARMA";
                    workSheet.Cells[1, 21].Value = "DIA AGRUPACION";

                    //Añadimos registros de España
                    while (r.Read())
                    {
                        f++;
                        c = 1;

                        #region BUSCAR DATOS EN T_ED_H_PS_HIST
                        //Si el NIF está vacío supuestamente es porque el contrato está de baja y sale de la tabla cont.t_ed_h_ps
                        //Intentamos recuperar los campos EMPRESA, TIPO_CLIENTE, NIF, CLIENTE, TARIFA, CONTRATO Y DISTRIBUIDORA del histórico, cont.t_ed_h_ps_hist
                        Dictionary<string, string> datos_histo = new Dictionary<string, string>();
                        string dat = "";
                        if (r["NIF"] == System.DBNull.Value && r["CUPS20"] != System.DBNull.Value)
                        {
                            strSql_histo = "SELECT ps.cd_empr AS EMPRESA, ps.de_tp_cli AS TIPO_CLIENTE, ps.cd_nif_cif_cli AS NIF, ps.tx_apell_cli AS CLIENTE, ps.cd_tarifa_c AS TARIFA, ps.cd_crto_comercial AS CONTRATO, ps.de_empr_distdora_nombre AS DISTRIBUIDORA "
                                + "FROM cont.t_ed_h_ps_pt_hist ps WHERE ps.cups20 = '" + r["CUPS20"] + "' ORDER BY ps.fh_act_dmco DESC LIMIT 1;";

                            db_histo = new MySQLDB(MySQLDB.Esquemas.GBL);
                            command_histo = new MySqlCommand(strSql_histo, db_histo.con);
                            r_histo = command_histo.ExecuteReader();
                            if (r_histo.Read())
                            {
                                if (r_histo["EMPRESA"] != System.DBNull.Value)
                                    datos_histo.Add("EMPRESA", r_histo["EMPRESA"].ToString());
                                if (r_histo["TIPO_CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("TIPO_CLIENTE", r_histo["TIPO_CLIENTE"].ToString());
                                if (r_histo["NIF"] != System.DBNull.Value)
                                    datos_histo.Add("NIF", r_histo["NIF"].ToString());
                                if (r_histo["CLIENTE"] != System.DBNull.Value)
                                    datos_histo.Add("CLIENTE", r_histo["CLIENTE"].ToString());
                                if (r_histo["TARIFA"] != System.DBNull.Value)
                                    datos_histo.Add("TARIFA", r_histo["TARIFA"].ToString());
                                if (r_histo["CONTRATO"] != System.DBNull.Value)
                                    datos_histo.Add("CONTRATO", r_histo["CONTRATO"].ToString());
                                if (r_histo["DISTRIBUIDORA"] != System.DBNull.Value)
                                    datos_histo.Add("DISTRIBUIDORA", r_histo["DISTRIBUIDORA"].ToString());
                            }

                            db_histo.CloseConnection();
                        }
                        #endregion

                        //Creamos objeto ResumenAgrupadaPendiente para ir almacenando info y al final, si procede, añadir a dic dic_resumen_agrupadas_ES
                        //EndesaEntity.facturacion.ResumenAgrupadaPendiente rap = new EndesaEntity.facturacion.ResumenAgrupadaPendiente();
                        //tiene_num_dia_agrupacion = false;


                        if (r["EMPRESA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["EMPRESA"].ToString();
                        else if (datos_histo.TryGetValue("EMPRESA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["SEGMENTO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["SEGMENTO"].ToString();
                        else
                            workSheet.Cells[f, c].Value = "NULO";
                        c++;

                        if (r["TIPO_CLIENTE"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TIPO_CLIENTE"].ToString();
                        else if (datos_histo.TryGetValue("TIPO_CLIENTE", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["NIF"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["NIF"].ToString();
                            //rap.cif = r["NIF"].ToString();
                        }
                        else if (datos_histo.TryGetValue("NIF", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            //rap.cif = dat;
                        }
                        c++;

                        if (r["CLIENTE"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["CLIENTE"].ToString();
                            //rap.razon_social = r["CLIENTE"].ToString();
                        }
                        else if (datos_histo.TryGetValue("CLIENTE", out dat))
                        {
                            workSheet.Cells[f, c].Value = dat;
                            //rap.razon_social = dat;
                        }
                        c++;

                        if (r["FALTACONT"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["FALTACONT"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["FPSERCON"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDateTime(r["FPSERCON"]).Date;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (r["CUPS20"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CUPS20"].ToString();
                        c++;

                        if (r["N_INSTALACION"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["N_INSTALACION"].ToString();
                        c++;

                        if (r["TARIFA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                        else if (datos_histo.TryGetValue("TARIFA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["CONTRATO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CONTRATO"].ToString();
                        else if (datos_histo.TryGetValue("CONTRATO", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["MES"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["MES"]);
                            aniomes = Convert.ToInt32(r["MES"]);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                        }
                        c++;

                        if (r["DISTRIBUIDORA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["DISTRIBUIDORA"].ToString();
                        else if (datos_histo.TryGetValue("DISTRIBUIDORA", out dat))
                            workSheet.Cells[f, c].Value = dat;
                        c++;

                        if (r["ESTADO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["ESTADO"].ToString();
                        c++;

                        if (r["SUBESTADO"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["SUBESTADO"].ToString();
                            //switch (r["SUBESTADO"].ToString().Substring(0, 2))
                            //{
                            //    case "01":
                            //        rap.pendiente_medida = 1;
                            //        break;
                            //    case "02":
                            //        rap.oc_calculable = 1;
                            //        break;
                            //    case "03":
                            //        rap.dc_generado_sin_di = 1;
                            //        break;
                            //    case "04":
                            //        rap.di_apartado = 1;
                            //        break;

                            //}
                        }
                        c++;

                        //if (r["lg_multimedida"] != System.DBNull.Value)
                        //{
                        //    workSheet.Cells[f, c].Value = r["lg_multimedida"].ToString();

                        //}
                        //else
                        //{
                        //    workSheet.Cells[f, c].Value = "N";
                        //}

                        workSheet.Cells[f, c].Value = GetDiasCUPSSubEstado(r["CUPS20"].ToString(), r["SUBESTADO"].ToString().Split(' ')[0]);

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["TAM"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            //rap.tam_total = rap.tam_total + Convert.ToDouble(r["TAM"]);
                        }
                        c++;

                        // NO MOSTRAMOS EN ESTE CASO LOS MESES NI IMPORTES PENDIENTES
                        //if (r["MES"] != System.DBNull.Value)
                        //{
                        //    //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                        //    meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                        //    workSheet.Cells[f, c].Value = meses_pdtes;
                        //    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //}
                        //c++;

                        //if (r["MES"] != System.DBNull.Value && r["TAM"] != System.DBNull.Value)
                        //{
                        //    //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                        //    meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                        //    workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]) * meses_pdtes;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        //}
                        //else
                        //{
                        //    c++;
                        //}

                        if (r["AGORA"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["AGORA"].ToString();
                            //if (r["AGORA"].ToString() == "S")
                            //    rap.es_agora = true;
                        }

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["TIPO_ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["TIPO_ALARMA"].ToString();

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                        if (r["ALARMA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["ALARMA"].ToString();

                        c++;

                        if (r["DIA_AGRUPACION"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = r["DIA_AGRUPACION"].ToString();
                            //rap.num_dia_agrupacion = Convert.ToInt16(r["DIA_AGRUPACION"]);
                            //tiene_num_dia_agrupacion = true;
                        }

                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        
                    }
                    db.CloseConnection();


                    #endregion


                    #endregion


                    headerCells = workSheet.Cells[1, 1, 1, c];
                    headerFont = headerCells.Style.Font;
                    headerFont.Bold = true;
                    allCells = workSheet.Cells[1, 1, f, c];

                    allCells.AutoFitColumns();
                    workSheet.Cells[1, 1, 1, c].AutoFilter = true;

                    

                    #endregion

                    //Guardarmos el informe estados descartados del pendiente SAP
                    excelPackage_estados_descartados.SaveAs(fileInfoEstadosDescartados);

                    if (automatico && param.GetValue("mail_enviar_mail_psat_tam") == "S")
                    {
                        ss_pp.Update_Fecha_Fin("Facturación", "Informe Pendiente BI", "Informe Pendiente BI");
                        EnvioCorreo_PdteWeb_BI(ruta_salida_archivo, ruta_salida_archivo_estados_descartados, fecha_informe);

                        //Añadimos ejecución procesos externos (PACO)

                        // 1. Lanzamos ejecutable Paco Incidencias SAP - 04/06/2024
                        ss_pp.Update_Fecha_Inicio("Facturación", "Informe Pendiente BI", "Informe Incidencias SAP");
                        //ss_pp.Update_Comentario("Facturación", "Informe Pendiente BI", "Informe Incidencias SAP", "Ejecutando Informe Incidencias SAP");
                        ficheroLog.Add("Ejecutando Informe Incidencias SAP: " + param.GetValue("ruta_ejecutable_informe_incidencias_sap"));
                        
                        utilidades.Fichero.EjecutaComando(param.GetValue("ruta_ejecutable_informe_incidencias_sap"), null);
                        
                        ficheroLog.Add("Fin ejecución Informe Incidencias SA: " + param.GetValue("ruta_ejecutable_informe_incidencias_sap"));
                        ss_pp.Update_Fecha_Fin("Facturación", "Informe Pendiente BI", "Informe Incidencias SAP");
                        //ss_pp.Update_Comentario("Facturación", "Informe Pendiente BI", "Informe Incidencias SAP", "OK");


                        // 2. Lanzamos Informe SAP-KRONOS - 04/07/2024
                        ficheroLog.Add("Ejecutando Informe SAP-KRONOS: EndesaBusiness.facturacion.redshift.Pendiente_Facturacion_Medida_B2B()");

                        EndesaBusiness.facturacion.redshift.Pendiente_Facturacion_Medida_B2B pendDos = new EndesaBusiness.facturacion.redshift.Pendiente_Facturacion_Medida_B2B();
                        pendDos.GeneraInformePendSAP(false);


                        //3. Lanzamos aplicativo Informe Pendiente Operaciones Eloy
                        ss_pp.Update_Fecha_Inicio("Facturación", "Informe Pendiente BI", "Informe Pendiente Operaciones Eloy");
                        //ss_pp.Update_Comentario("Facturación", "Informe Pendiente BI", "Informe Pendiente Operaciones Eloy", "Ejecutando Informe Pendiente Operaciones Eloy");
                        ficheroLog.Add("Ejecutando Informe Pendiente Operaciones Eloy: " + param.GetValue("ruta_ejecutable_informe_pdte_oper_eloy"));

                        utilidades.Fichero.EjecutaComando(param.GetValue("ruta_ejecutable_informe_pdte_oper_eloy"), null);

                        ficheroLog.Add("Fin ejecución Informe Pendiente Operaciones Eloy: " + param.GetValue("ruta_ejecutable_informe_pdte_oper_eloy"));
                        ss_pp.Update_Fecha_Fin("Facturación", "Informe Pendiente BI", "Informe Pendiente Operaciones Eloy");
                        //ss_pp.Update_Comentario("Facturación", "Informe Pendiente BI", "Informe Pendiente Operaciones Eloy", "OK");


                        //4. Lanzamos comando externo Informe SAP_Instalaciones
                        ss_pp.Update_Fecha_Inicio("Facturación", "Informe Pendiente BI", "Informe SAP_Instalaciones");
                        ficheroLog.Add("Ejecutando Informe SAP Instalaciones: " + param.GetValue("ruta_ejecutable_informe_sap_instalaciones"));

                        utilidades.Fichero.EjecutaComando(param.GetValue("ruta_ejecutable_informe_sap_instalaciones"), null);

                        ficheroLog.Add("Fin ejecución Informe SAP Instalaciones: " + param.GetValue("ruta_ejecutable_informe_sap_instalaciones"));
                        ss_pp.Update_Fecha_Fin("Facturación", "Informe Pendiente BI", "Informe SAP_Instalaciones");
                        
                        //Fin ejecución procesos externos

                    }

                    #region comentada
                    //Nuevo correo SAP. España y Portugal. Informe Pendiente. Subestados: 01.Z No definido y 03.Z Desconocido *********************
                    //*****************************************************************************************************************************

                    /*strSql = "SELECT ps.cd_empr as EMPRESA,   ps.cd_nif_cif_cli as NIF, ps.tx_apell_cli as CLIENTE,"
                       + " ps.cups20 as CUPS20, p.id_instalacion as N_INSTALACION, "
                       + " ps.cd_crto_comercial as CONTRATO, p.fh_periodo as MES,  "
                       + " concat(p.cd_estado,' ',de.de_estado) as ESTADO, concat(p.cd_subestado,' ',ds.de_subestado) as SUBESTADO, '' as DIAS_ESTADO, p.TAM, '' as MESES_PDTES_FACTURAR, '' as IMPORTE_PDTE_FACTURAR "
                       + " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p"
                       + " LEFT OUTER JOIN cont.t_ed_h_ps ps ON"
                       + " ps.cups20 = p.cd_cups"
                       + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                       + " de.cd_estado = p.cd_estado"
                       + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                       + " ds.cd_subestado = p.cd_subestado"
                       + " where p.fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "' and"
                       + " p.cd_empr_titular in ('ES21','ES22') and"
                       + " concat(p.cd_subestado,' ',ds.de_subestado) in ('01.Z No definido','03.Z Desconocido')";

                    db = new MySQLDB(MySQLDB.Esquemas.GBL);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    f = 1;

                    string archivoDesconocido;
                    archivoDesconocido = @"c:\Temp\NoDefinidos" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
                    FileInfo fileInfoNoDefinido = new FileInfo(archivoDesconocido);
                    excelPackage = new ExcelPackage(fileInfoNoDefinido);
                    //Creo la primera hoja
                    workSheet = excelPackage.Workbook.Worksheets.Add("Datos");
                    headerCells = workSheet.Cells[1, 1, 1, 17];
                    headerFont = headerCells.Style.Font;

                    //Exporta cabeceras de las columnas que queremos sacar   - no nos vale en este caso, hay muchos 
                    //alias y calculos de columas
                    for (int i = 0; i < r.FieldCount; i++)
                    {
                        workSheet.Cells[1, i + 1].Value = r.GetName(i);
                    }

                    while (r.Read())
                    {
                        f++;
                        c = 1;

                        if (r["EMPRESA"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["EMPRESA"].ToString();
                        c++;

                        if (r["NIF"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["NIF"].ToString();
                        c++;

                        if (r["CLIENTE"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CLIENTE"].ToString();
                        c++;

                        if (r["cups20"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["cups20"].ToString();
                        c++;

                        if (r["N_INSTALACION"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["N_INSTALACION"].ToString();
                        c++;

                        if (r["CONTRATO"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["CONTRATO"].ToString();
                        c++;

                        if (r["MES"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToInt32(r["MES"]);
                            aniomes = Convert.ToInt32(r["MES"]);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);
                        }
                        c++;


                        if (r["estado"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["estado"].ToString();
                        c++;

                        if (r["subestado"] != System.DBNull.Value)
                            workSheet.Cells[f, c].Value = r["subestado"].ToString();
                        c++;

                        workSheet.Cells[f, c].Value = GetDiasEstado(r["cups20"].ToString());


                        if (r["TAM"] != System.DBNull.Value)
                        {
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]);
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        }
                        c++;

                        if (r["MES"] != System.DBNull.Value)
                        {
                            //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                            workSheet.Cells[f, c].Value = meses_pdtes;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        c++;

                        if (r["MES"] != System.DBNull.Value && r["TAM"] != System.DBNull.Value)
                        {
                            //int meses_pdtes = mes_actual - Convert.ToInt32(r["mes"]);
                            meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                            workSheet.Cells[f, c].Value = Convert.ToDouble(r["TAM"]) * meses_pdtes;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        }
                        else
                        {
                            c++;
                        }


                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        c++;

                    }
                    db.CloseConnection();

                    allCells = workSheet.Cells[1, 1, f, c];
                    allCells.AutoFitColumns();
                    /*

                    //SpreadsheetGear.Drawing.Image image = new SpreadsheetGear.Drawing.Image(worksheet.Cells["A1:D10"]);

                    // Create Bitmap of range
                    //System.Drawing.Bitmap bitmap = image.GetBitmap();

                    ExcelRange Rango;
                    Rango = workSheet.Cells["$C$1:$C$3"];
                 
                    //Rango.CopyPicture (A xlScreen, xlBitmap);
                    //FileInfo prueba = new FileInfo("c:/Temp/Preuba.bmp");
                    //Rango.SaveToText(prueba,ExcelWorksheet);
                 

               
                   
                    /*
                     * 
                     *  rangoDeCeldas.Select();
                    workSheet.Cells["$A$1:$CJ$5"].Copy();
                    //Exportamos a imagen el rango
                    excelPackage.Chart MyChart  ;
                    string rango;
                    int indicador;
  
                    indicador = Ancho * 60

                    rango = "$A1$:$J$" & CStr(c++)

                    With excelPackage.Sheets(hoja).Range(rango)

                        System.Threading.Thread.Sleep(5000)
                        .select
                        .copy
                        .CopyPicture(Appearance:= excelPackage.XlPictureAppearance.xlScreen, Format:= excelPackage.XlCopyPictureFormat.xlPicture);
                        System.Threading.Thread.Sleep(5000);
     
                        'creo el objeto gráfico:  ancho y alto paso los valores
                        'que tiene el rango, así la imagen es del mismo tamaño (Left, Top, Width, Heigh)
                        If Archivo = "InformeCMM" Then  'ES para el informe CMM
                            MyChart = .Parent.ChartObjects.Add(10, 10, indicador, 950).Chart
                        ElseIf Archivo = "Foto" Then
                            MyChart = .Parent.ChartObjects.Add(10, 10, 350, 150).Chart
                        Else 'Es para el informe Diario Gas
                            MyChart = .Parent.ChartObjects.Add(10, 10, 900, 750).Chart
                        End If

                    End With

                    System.Threading.Thread.Sleep(2000)

                    //pego dentro del gráfico el rango copiado previamente
                    MyChart.ChartArea.Select()
                    MyChart.Paste()

                    //y le quito los bordes:
                    MyChart.ChartArea.Border.LineStyle = 0


                    System.Threading.Thread.Sleep(1000)

                    MyChart.Export(Filename:= strRutaPicture & Archivo & "_" & CStr(Format(Now, "ddMMyyyy")) & ".gif", FilterName:= "GIF")
                    System.Threading.Thread.Sleep(1000)
     
                    MyChart.Parent.Delete

                    

                    excelPackage.SaveAs(fileInfoNoDefinido);

                    

                    if (f > 1 && automatico) {
                        // La consulta ha devuelto datos, puedo envíar el correo
                        EnvioCorreo_PdteWeb_BI(archivoDesconocido, fecha_informe);
                    }
                    */
                    //**********************************************************************************
                    #endregion

                    if (automatico)
                        ss_pp.Update_Fecha_Fin("Facturación", "Informe Pendiente BI", "Informe Pendiente BI");
                }
                else if(automatico)
                {
                    ss_pp.Update_Comentario("Facturación", "Informe Pendiente BI", "Informe Pendiente BI",
                        "La fecha de actualización en BI es: " + UltimaActualizacionMySQL().Date.ToString("dd/MM/yyyy"));
                }
            }
        
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
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

            try
            {




                //sof.Contruye_Sofisticados();
                //agoraManual = CargaAgoraManual(DateTime.Now, DateTime.Now);
                //agoraPortugal = new contratacion.Agora_Portugal();

                //strSql = " SELECT fecha_informe, pend.empresa_titular AS EMPRESA,"
                //    + " pend.cups13, "
                //    + " pend.mes as aaaammPdte, pend.estado, pend.subestado, pend.tam"
                //    + " FROM fact. pend where "
                //    + " fecha_informe >= '" + f.ToString("yyyy-MM-dd") + "'"
                //    + " ORDER BY pend.fecha_informe, pend.empresa_titular, "
                //    + " pend.cups13, pend.mes ASC";

                strSql = "SELECT p.cd_empr_titular, ps.cd_empr, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                        + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, p.cd_cups as cups20, ps.cd_tarifa_c,"
                        + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre, p.cd_estado, p.cd_subestado,"
                        + " de.de_estado, ds.de_subestado, p.fh_periodo, IF(s.CUPS20 IS NULL, p.agora, 'S') AS agora, p.TAM,"
                        + " p.lg_multimedida, p.fec_act"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"
                        + " LEFT OUTER JOIN fact.cm_sofisticados s on"
                        + " s.CUPS20 = p.cd_cups"
                        + " where p.fec_act >= '" + DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd") + "'";

                foreach(string p in lista_empresas)
                {
                    if (firstOnly)
                    {
                        strSql += " and cd_empr_titular in ("
                            + "'" + p + "'";
                        firstOnly = false;
                    }else
                        strSql += ",'" + p + "'";
                }
                

                        
                   strSql += ") ORDER BY p.fec_act desc, ps.cd_empr, "
                        + " ps.cups20, p.fh_periodo ASC";


                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {                    
                    EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                    c.cod_empresaTitular = r["cd_empr_titular"].ToString();
                    c.empresaTitular = r["cd_empr"].ToString();                    
                    c.cups20 =  r["cups20"].ToString();
                    c.aaaammPdte = Convert.ToInt32(r["fh_periodo"]);
                    c.cod_estado = r["cd_estado"].ToString();
                    c.cod_subestado = r["cd_subestado"].ToString();
                    c.estado = r["de_estado"].ToString();
                    c.subsEstado = r["de_subestado"].ToString();
                    c.fecha_informe = Convert.ToDateTime(r["fec_act"]).Date;


                    if (r["tam"] != System.DBNull.Value)
                    {
                        if (c.aaaammPdte != 0)
                        {
                            aniomes = Convert.ToInt32(c.aaaammPdte);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);

                            // [21-01-2025 GUS] Modificamos cálculo de meses pendientes para que tenga en cuenta la fecha de actualización 'fec_act' en vez de la fecha actual de ejecución del informe
                            // y evitamos incoherencias en el cambio de mes
                            //meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                            meses_pdtes = ((c.fecha_informe.Year - fecha_registro.Year) * 12) + c.fecha_informe.Month - fecha_registro.Month;

                            c.tam = Convert.ToDouble(r["tam"]) * meses_pdtes;
                        }
                        else
                            c.tam = Convert.ToDouble(r["tam"]);

                    }

                    if (r["agora"] != System.DBNull.Value)
                        if(r["agora"].ToString() != "N")
                            c.agora = true;
                        else
                            c.agora = false;
                    else
                        c.agora = false;

                    if (r["lg_multimedida"] != System.DBNull.Value)
                        c.multimedida = r["lg_multimedida"].ToString() == "S";
                    else
                        c.multimedida = false;


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
                    

                   



                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                ficheroLog.addError("CargaPendiente: " + e.Message);
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
            bool isBTN = false;

            try
            {

                //sof.Contruye_Sofisticados();
                //agoraManual = CargaAgoraManual(DateTime.Now, DateTime.Now);
                //agoraPortugal = new contratacion.Agora_Portugal();
                

                strSql = "SELECT p.cd_empr_titular, ps.cd_empr, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                        + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, p.cd_cups, ps.cd_tarifa_c, p.cd_segmento_ptg,"
                        + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre, p.cd_estado, p.cd_subestado,"
                        + " de.de_estado, ds.de_subestado, p.fh_periodo, IF(s.CUPS20 IS NULL, p.agora, 'S') AS agora , p.TAM,"
                        + " p.lg_multimedida, p.fec_act"
                        + " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p"
                        + " LEFT OUTER JOIN cont.t_ed_h_ps_pt ps ON"
                        + " ps.cups20 = p.cd_cups"
                        + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                        + " de.cd_estado = p.cd_estado"
                        + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                        + " ds.cd_subestado = p.cd_subestado"
                        + " LEFT OUTER JOIN fact.cm_sofisticados s on"
                        + " s.CUPS20 = p.cd_cups"
                        + " where p.fec_act >= '" + DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd") + "'";

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
                if (!firstOnly)
                    strSql += ")";

                firstOnly = true;
                foreach (string p in lista_segmentos)
                {
                    if (p == "BTN")
                        isBTN = true;

                    if (firstOnly)
                    {
                        strSql += " and (cd_segmento_ptg in ("
                            + "'" + p + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + p + "'";
                }

                if (!firstOnly)
                {
                    if (isBTN) //Si la lista de segmentos incluye el segmento BTN incluimos tambien aquellos que no tienen informado el segmento (Petición de Ignacio - 26/02/2025)
                        strSql += ") or cd_segmento_ptg is null) ORDER BY p.fec_act desc, ps.cd_empr, "
                         + " ps.cups20, p.fh_periodo ASC";
                    else
                        strSql += ")) ORDER BY p.fec_act desc, ps.cd_empr, "
                         + " ps.cups20, p.fh_periodo ASC";
                }
                else
                {
                    strSql += " ORDER BY p.fec_act desc, ps.cd_empr, "
                         + " ps.cups20, p.fh_periodo ASC";                     
                }



                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
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

                    if (r["cd_segmento_ptg"] != System.DBNull.Value)
                        c.segmento = r["cd_segmento_ptg"].ToString();
                    else
                        c.segmento = "NULL";

                    if (r["tam"] != System.DBNull.Value)
                    {
                        if (c.aaaammPdte != 0)
                        {
                            aniomes = Convert.ToInt32(c.aaaammPdte);
                            fecha_registro = new DateTime(Convert.ToInt32(aniomes.ToString().Substring(0, 4)),
                                Convert.ToInt32(aniomes.ToString().Substring(4, 2)), 1);

                            // [21-01-2025 GUS] Modificamos cálculo de meses pendientes para que tenga en cuenta la fecha de actualización 'fec_act' en vez de la fecha actual de ejecución del informe
                            // y evitamos incoherencias en el cambio de mes
                            //meses_pdtes = ((fecha_actual.Year - fecha_registro.Year) * 12) + fecha_actual.Month - fecha_registro.Month;
                            meses_pdtes = ((c.fecha_informe.Year - fecha_registro.Year) * 12) + c.fecha_informe.Month - fecha_registro.Month;                            
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

                    //if (r["lg_multimedida"] != System.DBNull.Value)
                    //    c.multimedida = r["lg_multimedida"].ToString() == "S";
                    //else
                    //    c.multimedida = false;



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
                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                ficheroLog.addError("CargaPendiente: " + e.Message);
                return null;
            }
        }

        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaTotales(List<string> lista_empresas)
        {
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>>();




            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string estado = "";
            string subestado = "";
            int total_cups = 0;
            double total_tam = 0;
            DateTime fecha_informe = new DateTime();

            try
            {
                strSql = "SELECT t.fh_envio, t.cd_estado, t.cd_subestado, t.num_cups, t.tam"
                    + " FROM t_ed_h_sap_pendiente_facturar_agrupado_totales t"
                    + " where cd_empr_titular in ('" + lista_empresas[0] + "'";

                for (int x = 1; x < lista_empresas.Count; x++)
                    strSql += ",'" + lista_empresas[x] + "'";

                strSql += ") ORDER BY t.fh_envio DESC, t.cd_estado";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha_informe = Convert.ToDateTime(r["fh_envio"]);
                    estado = r["cd_estado"].ToString();
                    subestado = r["cd_subestado"].ToString();
                    total_cups = Convert.ToInt32(r["num_cups"]);

                    if (r["tam"] != System.DBNull.Value)
                        total_tam = Convert.ToDouble(r["tam"]);

                    List<EndesaEntity.medida.Pendiente_Totales> o;
                    if (!d.TryGetValue(fecha_informe, out o))
                    {
                        o = InicializaPendienteTotales();
                        d.Add(fecha_informe, o);
                    }


                    foreach (EndesaEntity.medida.Pendiente_Totales p in o)
                    {
                        if (p.estado == estado && p.subestado == subestado)
                        {
                            p.num_cups += total_cups;
                            p.tam += total_tam;
                        }

                    }


                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {

                return null;

            }
        }

        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaAgora_TAM(DateTime fd, List<string> lista_empresas)
        {
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>>();




            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string estado = "";
            string subestado = "";
            int total_cups = 0;
            double tam = 0;
            DateTime fecha_informe = new DateTime();
            utilidades.Fechas utilFechas = new Fechas();

            try
            {
                strSql = "SELECT t.fh_envio, t.cd_estado, t.cd_subestado, t.num_cups, t.tam"
                     + " FROM t_ed_h_sap_pendiente_facturar_agrupado_totales t"
                     + " where cd_empr_titular in ('" + lista_empresas[0] + "'";

                for (int x = 1; x < lista_empresas.Count; x++)
                    strSql += ",'" + lista_empresas[x] + "'";

                strSql += ") and t.agora = 'S'" 
                    + " and t.fh_envio >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " ORDER BY t.fh_envio, t.cd_estado";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    fecha_informe = Convert.ToDateTime(r["fh_envio"]);
                    estado = r["cd_estado"].ToString();
                    subestado = r["cd_subestado"].ToString();
                    total_cups = Convert.ToInt32(r["num_cups"]);

                    if (r["tam"] != System.DBNull.Value)
                        tam = Convert.ToDouble(r["tam"]);

                    List<EndesaEntity.medida.Pendiente_Totales> o;
                    if (!d.TryGetValue(fecha_informe, out o))
                    {
                        o = InicializaPendienteTotales();
                        d.Add(fecha_informe, o);
                    }

                    foreach (EndesaEntity.medida.Pendiente_Totales p in o)
                    {
                        if (p.estado == estado && p.subestado == subestado)
                        {
                            p.num_cups += total_cups;
                            p.tam += tam;
                        }

                    }


                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {

                return null;

            }
        }

        private Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> CargaNoAgora_TAM(DateTime fd, List<string> lista_empresas)
        {
            Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>> d
                = new Dictionary<DateTime, List<EndesaEntity.medida.Pendiente_Totales>>();

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string estado = "";
            string subestado = "";
            int total_cups = 0;
            double tam = 0;
            DateTime fecha_informe = new DateTime();
            utilidades.Fechas utilFechas = new Fechas();

            try
            {
                strSql = "SELECT t.fh_envio, t.cd_estado, t.cd_subestado, t.num_cups, t.tam"
                      + " FROM t_ed_h_sap_pendiente_facturar_agrupado_totales t"
                      + " where cd_empr_titular in ('" + lista_empresas[0] + "'";

                for (int x = 1; x < lista_empresas.Count; x++)
                    strSql += ",'" + lista_empresas[x] + "'";

                strSql += ") and t.agora = 'N'"
                    + " and t.fh_envio >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " ORDER BY t.fh_envio, t.cd_estado";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha_informe = Convert.ToDateTime(r["fh_envio"]);
                    estado = r["cd_estado"].ToString();
                    subestado = r["cd_subestado"].ToString();
                    total_cups = Convert.ToInt32(r["num_cups"]);

                    if (r["tam"] != System.DBNull.Value)
                        tam = Convert.ToDouble(r["tam"]);

                    List<EndesaEntity.medida.Pendiente_Totales> o;
                    if (!d.TryGetValue(fecha_informe, out o))
                    {
                        o = InicializaPendienteTotales();
                        d.Add(fecha_informe, o);
                    }


                    foreach (EndesaEntity.medida.Pendiente_Totales p in o)
                    {
                        if (p.estado == estado && p.subestado == subestado)
                        {
                            p.num_cups = total_cups;
                            p.tam = tam;
                        }

                    }


                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {

                return null;

            }
        }

        private List<EndesaEntity.medida.Pendiente_Totales> InicializaPendienteTotales()
        {
            List<EndesaEntity.medida.Pendiente_Totales> t = new List<EndesaEntity.medida.Pendiente_Totales>();
            

            Pendiente_Estados estados = new Pendiente_Estados();
            Pendiente_Subestados subestados = new Pendiente_Subestados();


            foreach(KeyValuePair<string, EndesaEntity.medida.Pendiente> p in subestados.dic)
            {
                string[] texto = p.Key.Split('.');

                EndesaEntity.medida.Pendiente_Totales c = new EndesaEntity.medida.Pendiente_Totales();
                c.estado = texto[0];
                c.subestado = p.Key;
                t.Add(c);
            }
                       

            return t;
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
                            if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora &&
                                (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
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
                                if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z]
                                    && o[i].agora == agora &&
                                    (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
                                    total = total + 1;

                            }
                }
            }                

            return total;

        }
              


        private double Total_Pendiente_TAM(bool agora, DateTime fecha, List<string> lista_empresas, List<string> lista_segmentos, string estado, string subestado)
        {
            double total = 0;

            List<EndesaEntity.medida.Pendiente> o;

            if(lista_segmentos == null)
            {
                if (dic_pendiente_hist_fecha.TryGetValue(fecha, out o))
                {
                    for (int j = 0; j < lista_empresas.Count; j++)
                        for (int i = 0; i < o.Count; i++)
                        {
                            if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].agora == agora &&
                                (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
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
                            if (o[i].cod_empresaTitular == lista_empresas[j] && o[i].segmento == lista_segmentos[z]
                                    && o[i].agora == agora &&
                                (o[i].cod_estado == estado && (o[i].cod_subestado == subestado || subestado == null)))
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
            
            strSql = "SELECT ps.cd_empr, ps.cd_tp_tension, ps.cd_nif_cif_cli, ps.de_tp_cli, ps.tx_apell_cli,"
                + " ps.fh_alta_crto, ps.fh_inicio_vers_crto, p.cd_cups as cups20, p.id_instalacion, ps.cd_tarifa_c,"
                + " ps.cd_crto_comercial, ps.de_empr_distdora_nombre, p.lg_multimedida,"
                + " concat(p.cd_estado,' ',de.de_estado) as de_estado, concat(p.cd_subestado,' ',if (ds.de_subestado is null,'', ds.de_subestado)) as de_subestado, p.fh_periodo as mes, p.agora, p.TAM"
                + " FROM fact.t_ed_h_sap_pendiente_facturar_agrupado p";

            if (lista_empresas[0].Contains("PT"))
                strSql += " LEFT OUTER JOIN cont.t_ed_h_ps_pt ps ON";
            else
                strSql += " LEFT OUTER JOIN cont.t_ed_h_ps ps ON";

            strSql += " ps.cups20 = p.cd_cups"
                + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de on"
                + " de.cd_estado = p.cd_estado"
                + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds on"
                + " ds.cd_subestado = p.cd_subestado"
                + " where p.fh_envio = '" + fecha_informe.Date.ToString("yyyy-MM-dd") + "' and"
                + " p.cd_empr_titular in (";
            
            foreach(string p in lista_empresas)
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

            if(lista_tension != null)
            {
                firstOnly = true;
                strSql += " and ps.cd_tp_tension in (";
                foreach (string p in lista_tension)
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
            }

            return strSql;

        }

        private void EnvioCorreo_PdteWeb_BI(string archivo,string archivo_estados_descartados, DateTime fecha_informe)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            FileInfo fileInfoEstadosDescartados = new FileInfo(archivo_estados_descartados);
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
                textBody.Append("  Se adjuntan los archivos ").Append(fileInfo.Name).Append(" y ").Append(fileInfoEstadosDescartados.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");

                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (param.GetValue("mail_enviar_mail_psat_tam") == "S")
                    mes.SendMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo + ";" + archivo_estados_descartados);

                else
                    mes.SaveMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo + ";" + archivo_estados_descartados);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("mail_from"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }

        private DateTime CalculaFechaDesdeinforme()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime date = DateTime.Now;

            strSql = "SELECT c.fh_envio"
                + " FROM t_ed_h_sap_pendiente_facturar_agrupado c"
                + " ORDER BY c.fh_envio desc"
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

        // Carga diccionario con la primera fecha que aparece el CUPS pendiente de facturar con código subestado actual, solo de los estados 03.C y 03.I
        // La clave del diccionario la formamos mediante concatenación de CUPS20 + cd_subestado
        private Dictionary<string, DateTime> CargaDiasCUPSSubEstado()
        {
            Dictionary<string, DateTime> d = new Dictionary<string, DateTime>();

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string cups = "";
            DateTime primera_fecha = DateTime.Now;

            //strSql = "SELECT p.cd_cups, p.fh_periodo, MIN(p.fh_envio) AS primera_fecha, p.cd_subestado"
            //    + " FROM t_ed_h_sap_pendiente_facturar p"
            //    + " WHERE p.fh_periodo >= " + DateTime.Now.AddYears(-1).ToString("yyyyMM")
            //    + " GROUP BY p.cd_cups, p.fh_periodo, p.cd_subestado"
            //    + " ORDER BY p.fh_envio DESC";

            strSql = "SELECT left(p.cd_cups, 20) AS cups20, p.cd_subestado, MIN(p.fh_envio) AS primera_fecha"
                + " FROM fact.t_ed_h_sap_pendiente_facturar p"
                + " INNER JOIN"

                + " (SELECT cd_cups, cd_subestado, fh_periodo"
                + " FROM fact.t_ed_h_sap_pendiente_facturar"
                + " WHERE cd_subestado IN ('03.I','03.C') AND fh_envio = (SELECT MAX(fh_envio) FROM fact.t_ed_h_sap_pendiente_facturar)"
                + " GROUP BY cd_cups, cd_subestado) AS pa"

                + " ON pa.cd_cups = p.cd_cups AND pa.cd_subestado = p.cd_subestado AND pa.fh_periodo = p.fh_periodo"
                + " GROUP BY p.cd_cups, p.cd_subestado;";


            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                cups = r["cups20"].ToString() + r["cd_subestado"].ToString();
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
        private int GetDiasCUPSSubEstado(string cups, string subestado)
        {
            int dias = 1;
            DateTime o;
            if (dic_dias_cups_subestado.TryGetValue(cups+subestado, out o))
            {
                dias = (DateTime.Now.Date - o.Date).Days + 1;
            }

            return dias;
        }

        private Dictionary<string, EndesaEntity.medida.Pendiente> Carga()
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";


            Dictionary<string, EndesaEntity.medida.Pendiente> d =
                new Dictionary<string, EndesaEntity.medida.Pendiente>();


            try
            {
                strSql = "SELECT p.cd_cups,"
                    + " concat(p.cd_estado, ' ', e.de_estado) as de_estado,"
                    + " concat(p.cd_subestado, ' ', s.de_subestado) as de_subestado,"
                    + " p.fh_periodo as mes"
                    + " FROM t_ed_h_sap_pendiente_facturar_agrupado p"
                    + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar e ON"
                    + " e.cd_estado = p.cd_estado"
                    + " LEFT OUTER JOIN t_ed_p_subestado_sap_pendiente_facturar s ON"
                    + " s.cd_subestado = p.cd_subestado"
                    + " WHERE p.fh_envio = (SELECT MAX(fh_envio) AS max_fh_envio FROM t_ed_h_sap_pendiente_facturar_agrupado)";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                    if (r["cd_cups"] != System.DBNull.Value)
                        c.cups20 = r["cd_cups"].ToString();

                    if (r["de_estado"] != System.DBNull.Value)
                        c.descripcion_estado = r["de_estado"].ToString();

                    if (r["de_subestado"] != System.DBNull.Value)
                        c.descripcion_subestado = r["de_subestado"].ToString();

                    if (r["mes"] != System.DBNull.Value)
                        c.aaaammPdte = Convert.ToInt32(r["mes"]);

                    EndesaEntity.medida.Pendiente o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception ex)
            {
                return null;
            }


        }

        public void GetEstados(string cups20)
        {
            EndesaEntity.medida.Pendiente o;
            if (dic_pendiente.TryGetValue(cups20, out o))
            {
                this.existe = true;
                this.cups20 = o.cups20;
                this.descripcion_estado = o.descripcion_estado;
                this.descripcion_subestado = o.descripcion_subestado;
                this.aaaammPdte = o.aaaammPdte;
            }
            else
                this.existe = false;
        }
    }
}

