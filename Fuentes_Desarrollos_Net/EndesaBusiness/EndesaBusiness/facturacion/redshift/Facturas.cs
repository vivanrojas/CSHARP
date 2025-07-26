using Aspose.Cells.Drawing;
using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using Microsoft.Graph;
using Microsoft.Win32;
using MySql.Data.MySqlClient;
using Renci.SshNet.Common;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static EndesaBusiness.medida.Kee_Extraccion_Formulas;

namespace EndesaBusiness.facturacion.redshift
{
    public class Facturas
    {

        utilidades.Seguimiento_Procesos ss_pp;
        logs.Log ficheroLog;
        logs.Log ficheroLogB2B;
        utilidades.Param p;

        public Facturas()
        {
            ss_pp = new utilidades.Seguimiento_Procesos();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Copia_Facturas_BI");
            ficheroLogB2B = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Copiar_Facturas_B2B");
            p = new utilidades.Param("copia_facturas_bi_param", MySQLDB.Esquemas.FAC);
        }


        public void CopiadoFacturasBI()
        {
            CopiadoFacturasCabecera();
            CopiadoFacturasConceptosBI();
        }
        
        public void CopiadoFacturasCabecera()
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

            try
            {

                ss_pp.Update_Fecha_Inicio("Facturación", "Copia Facturas BI", "Copia Facturas BI");

                ultimaFechaCopiado = UltimaActualizacion();

                ficheroLog.Add(Consulta(ultimaFechaCopiado));
                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(Consulta(ultimaFechaCopiado), db.con); 
                r = command.ExecuteReader();
                
                while (r.Read())
                {
                    j++;
                    k++;

                    if (firstOnly)
                    {
                        sb = null;
                        sb = new StringBuilder();
                        sb.Append("REPLACE INTO t_ed_h_sap_facts");
                        sb.Append(" (cl_empr, cd_mes, id_fact, cd_di, fh_fact, fh_ini_fact, fh_fin_fact,");
                        sb.Append(" cd_tp_fact, cd_est_fact, cd_tp_cli, nm_ener_consmda, nm_ener_factda,");
                        sb.Append(" im_factdo_sin_iva, im_factdo_con_iva, cl_cli, id_fact_anulada,");
                        sb.Append(" cd_di_sustyente, cd_di_sustituida, cd_di_anuladora, cd_cups_ext,");
                        sb.Append(" cd_cups_gas_ext, de_tp_fact, de_estado_fact, cd_cuenta_contr, cd_ind_imp_1, de_ind_imp_1,");
                        sb.Append(" cd_ind_imp_2, de_ind_imp_2, cd_ind_imp_3, de_ind_imp_3, im_base_imp1,");
                        sb.Append(" im_base_imp2, im_base_imp3, im_costes_atr, im_impuesto_1, im_impuesto_2,");
                        sb.Append(" im_impuesto_3, nm_total_igic, nm_total_ipsi, nm_total_iva, fec_act,");
                        sb.Append(" created_by, created_date) values ");
                        firstOnly = false;
                    }
                    #region Campos
                    
                    if (r["cl_empr"] != System.DBNull.Value)
                        sb.Append("('").Append(r["cl_empr"].ToString()).Append("',");
                    else
                        sb.Append("(null,");

                    if (r["cd_mes"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_mes"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_di"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_di"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_fact"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_ini_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_ini_fact"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_fin_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_fin_fact"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_est_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_est_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_cli"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_cli"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_ener_consmda"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_ener_consmda"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_ener_factda"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_ener_factda"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_factdo_sin_iva"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_factdo_sin_iva"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_factdo_con_iva"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_factdo_con_iva"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cl_cli"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cl_cli"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_fact_anulada"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_fact_anulada"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_di_sustyente"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_di_sustyente"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_di_sustituida"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_di_sustituida"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_di_anuladora"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_di_anuladora"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_cups_ext"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_cups_ext"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_cups_gas_ext"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_cups_gas_ext"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_estado_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_estado_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_cuenta_contr"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_cuenta_contr"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_ind_imp_1"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_ind_imp_1"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_ind_imp_1"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_ind_imp_1"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_ind_imp_2"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_ind_imp_2"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_ind_imp_2"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_ind_imp_2"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_ind_imp_3"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_ind_imp_3"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_ind_imp_3"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_ind_imp_3"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["im_base_imp1"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_base_imp1"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_base_imp2"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_base_imp2"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_base_imp3"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_base_imp3"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_costes_atr"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_costes_atr"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_impuesto_1"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_impuesto_1"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_impuesto_2"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_impuesto_2"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["im_impuesto_3"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["im_impuesto_3"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_total_igic"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_total_igic"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_total_ipsi"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_total_ipsi"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_total_iva"] != System.DBNull.Value)
                        sb.Append(Convert.ToDouble(r["nm_total_iva"]).ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fec_act"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fec_act"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");
                    #endregion

                    if (j == 50)
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



                ss_pp.Update_Fecha_Fin("Facturación", "Copia Facturas BI", "Copia Facturas BI");

            }
            catch(Exception ex)
            {
                ficheroLog.addError(ex.Message);
                ss_pp.Update_Fecha_Fin("Facturación", "Copia Facturas BI", "Copia Facturas BI");
                ss_pp.Update_Comentario("Facturación", "Copia Facturas BI", "Copia Facturas BI", "Error en CopiadoFacturasCabecera(): " + ex.Message);
            }
        }

        public void CopiadoFacturasConceptosBI()
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

            try
            {

                ultimaFechaCopiado = UltimaActualizacion();

                ficheroLog.Add(Consulta_Conceptos(ultimaFechaCopiado));
                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(Consulta_Conceptos(ultimaFechaCopiado), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    j++;
                    k++;

                    if (firstOnly)
                    {
                        sb = null;
                        sb = new StringBuilder();
                        sb.Append("REPLACE INTO t_ed_h_sap_facts_conceptos");
                        sb.Append(" (cd_di, cd_concepto, de_concepto, im_concepto,");                        
                        sb.Append(" created_by, created_date) values ");
                        firstOnly = false;
                    }
                    #region Campos


                    for (int i = 1; i <= 50; i++)
                    {
                        if (r["cd_concepto_" + i] != System.DBNull.Value)
                            if (r["cd_concepto_" + i].ToString() != "")
                            {
                            sb.Append("('").Append(r["cd_di"].ToString()).Append("',");
                            sb.Append("'").Append(r["cd_concepto_" + i].ToString()).Append("',");
                            sb.Append("'").Append(r["de_concepto_" + i].ToString()).Append("',");
                            sb.Append(r["im_concepto_" + i].ToString().Replace(",", ".")).Append(",");
                            sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                            sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");
                            }

                    }

                    if (j == 50)
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



                    #endregion



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

            }
            catch(Exception ex)
            {
                ficheroLog.addError(ex.Message);
            }
        }

        public DateTime UltimaActualizacion()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime fecha = new DateTime(2022,01,01);

            strSql = "SELECT max(fec_act) AS fec_act FROM t_ed_h_sap_facts";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
                fecha = Convert.ToDateTime(r["fec_act"]);
            db.CloseConnection();
            Console.WriteLine("Última fecha de copiado RedShift: " + fecha.ToString("dd/MM/yyyy"));
            ficheroLog.Add("Última fecha de copiado RedShift: " + fecha.ToString("dd/MM/yyyy"));

            return fecha;
        }

        public string Consulta(DateTime fecha)
        {
            string strSql = "";

            strSql = "SELECT f.cl_empr, f.cd_mes, f.id_fact, f.cd_di, f.fh_fact, f.fh_ini_fact, f.fh_fin_fact, f.cd_tp_fact, f.cd_est_fact, f.cd_tp_cli, f.nm_ener_consmda, f.nm_ener_factda," 
                + " f.im_factdo_sin_iva, f.im_factdo_con_iva, f.cl_cli, id_fact_anulada, f.cd_di_sustyente, f.cd_di_sustituida, f.cd_di_anuladora, cal.cd_cups_ext,"
                + " cal.cd_cups_gas_ext, f.de_tp_fact, f.de_estado_fact, f.cd_cuenta_contr, f.cd_ind_imp_1, f.de_ind_imp_1, f.cd_ind_imp_2, f.de_ind_imp_2, f.cd_ind_imp_3, f.de_ind_imp_3, f.im_base_imp1,"
                + " f.im_base_imp2, f.im_base_imp3, f.im_costes_atr, f.im_impuesto_1, f.im_impuesto_2, f.im_impuesto_3, f.nm_total_igic, f.nm_total_ipsi, f.nm_total_iva, f.fec_act"
                + " FROM ed_owner.t_ed_h_sap_facts f"
                + " inner join ed_owner.t_ed_h_sap_dc cal on"                
                + " cal.cd_di = f.cd_di where"
                + " f.fec_act >= '" + fecha.AddDays(-1000).ToString("yyyy-MM-dd") + "' AND"
                + " f.de_marca_back = 'OPERACIONES B2B'";

            return strSql;
        }

        public string Consulta_Conceptos(DateTime fecha)
        {
            string strSql = "";

            strSql = "SELECT cd_di, cd_concepto_1";

            for (int i = 2; i <= 50; i++)
                strSql += " ,cd_concepto_" + i;

            for (int i = 1; i <= 50; i++)
                strSql += " ,de_concepto_" + i;

            for (int i = 1; i <= 50; i++)
                strSql += " ,im_concepto_" + i;

            strSql +=  " FROM ed_owner.t_ed_h_sap_facts f where"
                + " f.fec_act >= '" + fecha.AddDays(-1000).ToString("yyyy-MM-dd") + "' AND"
                + " f.de_marca_back = 'OPERACIONES B2B'";

            return strSql;
        }

        //Función que lanza las copias de las facturas de B2B, los documentos de cálculo (CUPS) y los conceptos desde ed_owner al esquema de facturación facturacionb2b_owner
        public void CopiarFacturasB2B()
        {
            CopiarFacturasB2B_t_ed_h_sap_facts();
            CopiarFacturasB2B_t_ed_h_sap_facts_conceptos();
            CopiarFacturasB2B_t_ed_h_sap_dc();
        }
        public void CopiarFacturasB2B_t_ed_h_sap_facts()
        {
            StringBuilder sb = new StringBuilder();
            servidores.RedShiftServer db_source;
            servidores.RedShiftServer db_target;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";

            DateTime hora_inicio;
            DateTime hora_fin;
            TimeSpan diferencia_horas;
            int segundos;
            int segundos_totales=0;
            float reg_seg;
            int num_reg_insert = Convert.ToInt32(p.GetValue("num_reg_insert"));
            int minutos_restantes;

            bool firstOnly = true;
            int j = 0;
            int k = 0;
            int totalRegistros = 0;
            int totalRegistrosEliminados = 0;

            //Ejecutamos solo si NO se ha ejecutado correctamente hoy ==> NOT( Fecha Fin == HOY && Comentario == "OK") y NO ESTA EN EJECUCION ACTUALMENTE
            Console.WriteLine("Fecha fin proceso Copiar Facturas B2B - DI: " + ss_pp.GetFecha_FinProceso("Facturación", "Copiar Facturas B2B", "Copiar Facturas B2B - DI").Date.ToString());
            if (
                !((ss_pp.GetFecha_FinProceso("Facturación", "Copiar Facturas B2B", "Copiar Facturas B2B - DI").Date == DateTime.Now.Date) && (ss_pp.GetComentarioProceso("Facturación", "Copiar Facturas B2B", "Copiar Facturas B2B - DI") == "OK"))
                &&
                !(ss_pp.GetComentarioProceso("Facturación", "Copiar Facturas B2B", "Copiar Facturas B2B - DI") == "En ejecución")
                )
            {
                try
                {

                    ss_pp.Update_Fecha_Inicio("Facturación", "Copiar Facturas B2B", "Copiar Facturas B2B - DI");

                    db_source = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                    db_target = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD_FACTURACION);

                    //Obtenemos el total de registros
                    strSql = Total_Consulta_FacturasB2B_t_ed_h_sap_facts();
                    ficheroLogB2B.Add(strSql);
                    Console.WriteLine("\nEjecución consulta SQL [TOTAL FACTURAS - DI]: " + strSql);
                    command = new OdbcCommand(strSql, db_source.con);
                    r = command.ExecuteReader();
                    if (r != null)
                    {
                        r.Read();
                        totalRegistros = Convert.ToInt32(r["total"]);
                        r.Close();

                        //DELETE - Antes de insertar los registros borramos todos los registros de la tabla
                        #region BORRAR TABLA
                        strSql = Consulta_Delete_FacturasB2B_t_ed_h_sap_facts();
                        ficheroLogB2B.Add(strSql);
                        Console.WriteLine("\nEjecución consulta SQL [DELETE DETALLES FACTURAS - DI]: " + strSql);

                        command = new OdbcCommand(strSql, db_target.con);
                        totalRegistrosEliminados = command.ExecuteNonQuery();
                        ficheroLogB2B.Add("Eliminamos todos los registros de la tabla facturacionb2b_owner.t_ed_h_sap_facts: " + totalRegistrosEliminados + " registros eliminados.");
                        Console.WriteLine("\nEliminamos todos los registros de la tabla facturacionb2b_owner.t_ed_h_sap_facts: " + totalRegistrosEliminados + " registros eliminados.\n");
                        #endregion

                        //INSERT - Iniciamos anexión de datos
                        #region ANEXAR DATOS
                        strSql = Consulta_FacturasB2B_t_ed_h_sap_facts();
                        ficheroLogB2B.Add(strSql);
                        Console.WriteLine("\nEjecución consulta SQL [DETALLES FACTURAS - DI]: " + strSql);

                        command = new OdbcCommand(strSql, db_source.con);
                        //command.CommandText = strSql;
                        r = command.ExecuteReader();

                        ficheroLogB2B.Add("Iniciamos anexión de datos a " + num_reg_insert + " registros por consulta.");
                        Console.WriteLine("\nIniciamos anexión de datos a " + num_reg_insert + " registros por consulta.\n");
                        while (r.Read())
                        {
                            j++;
                            k++;

                            if (firstOnly)
                            {
                                sb = null;
                                sb = new StringBuilder();
                                sb.Append("INSERT INTO facturacionb2b_owner.t_ed_h_sap_facts");
                                sb.Append(" (cl_empr, cd_mes, id_fact, cd_di, fh_fact, fh_ini_fact, fh_fin_fact,");
                                sb.Append(" cd_tp_fact, cd_est_fact, cd_tp_cli, nm_ener_consmda, nm_ener_factda,");
                                sb.Append(" im_factdo_sin_iva, im_factdo_con_iva, cl_cli, id_fact_anulada,");
                                sb.Append(" cd_di_sustyente, cd_di_sustituida, cd_di_anuladora,");
                                sb.Append(" de_tp_fact, de_estado_fact, cd_cuenta_contr, cd_ind_imp_1, de_ind_imp_1,");
                                sb.Append(" cd_ind_imp_2, de_ind_imp_2, cd_ind_imp_3, de_ind_imp_3, im_base_imp1,");
                                sb.Append(" im_base_imp2, im_base_imp3, im_costes_atr, im_impuesto_1, im_impuesto_2,");
                                sb.Append(" im_impuesto_3, nm_total_igic, nm_total_ipsi, nm_total_iva, fec_act,");
                                sb.Append(" created_update_by, created_update_date) values ");
                                firstOnly = false;
                            }
                            #region Campos

                            if (r["cl_empr"] != System.DBNull.Value)
                                sb.Append("('").Append(r["cl_empr"].ToString()).Append("',");
                            else
                                sb.Append("(null,");

                            if (r["cd_mes"] != System.DBNull.Value)
                                sb.Append("'").Append(r["cd_mes"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["id_fact"] != System.DBNull.Value)
                                sb.Append("'").Append(r["id_fact"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["cd_di"] != System.DBNull.Value)
                                sb.Append("'").Append(r["cd_di"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["fh_fact"] != System.DBNull.Value)
                                sb.Append("'").Append(Convert.ToDateTime(r["fh_fact"]).ToString("yyyy-MM-dd")).Append("',");
                            else
                                sb.Append("null,");

                            if (r["fh_ini_fact"] != System.DBNull.Value)
                                sb.Append("'").Append(Convert.ToDateTime(r["fh_ini_fact"]).ToString("yyyy-MM-dd")).Append("',");
                            else
                                sb.Append("null,");

                            if (r["fh_fin_fact"] != System.DBNull.Value)
                                sb.Append("'").Append(Convert.ToDateTime(r["fh_fin_fact"]).ToString("yyyy-MM-dd")).Append("',");
                            else
                                sb.Append("null,");

                            if (r["cd_tp_fact"] != System.DBNull.Value)
                                sb.Append("'").Append(r["cd_tp_fact"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["cd_est_fact"] != System.DBNull.Value)
                                sb.Append("'").Append(r["cd_est_fact"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["cd_tp_cli"] != System.DBNull.Value)
                                sb.Append("'").Append(r["cd_tp_cli"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["nm_ener_consmda"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["nm_ener_consmda"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["nm_ener_factda"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["nm_ener_factda"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["im_factdo_sin_iva"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["im_factdo_sin_iva"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["im_factdo_con_iva"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["im_factdo_con_iva"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["cl_cli"] != System.DBNull.Value)
                                sb.Append("'").Append(r["cl_cli"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["id_fact_anulada"] != System.DBNull.Value)
                                sb.Append("'").Append(r["id_fact_anulada"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["cd_di_sustyente"] != System.DBNull.Value)
                                sb.Append("'").Append(r["cd_di_sustyente"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["cd_di_sustituida"] != System.DBNull.Value)
                                sb.Append("'").Append(r["cd_di_sustituida"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["cd_di_anuladora"] != System.DBNull.Value)
                                sb.Append("'").Append(r["cd_di_anuladora"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["de_tp_fact"] != System.DBNull.Value)
                                sb.Append("'").Append(r["de_tp_fact"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["de_estado_fact"] != System.DBNull.Value)
                                sb.Append("'").Append(r["de_estado_fact"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["cd_cuenta_contr"] != System.DBNull.Value)
                                sb.Append("'").Append(r["cd_cuenta_contr"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["cd_ind_imp_1"] != System.DBNull.Value)
                                sb.Append("'").Append(r["cd_ind_imp_1"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["de_ind_imp_1"] != System.DBNull.Value)
                                sb.Append("'").Append(r["de_ind_imp_1"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["cd_ind_imp_2"] != System.DBNull.Value)
                                sb.Append("'").Append(r["cd_ind_imp_2"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["de_ind_imp_2"] != System.DBNull.Value)
                                sb.Append("'").Append(r["de_ind_imp_2"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["cd_ind_imp_3"] != System.DBNull.Value)
                                sb.Append("'").Append(r["cd_ind_imp_3"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["de_ind_imp_3"] != System.DBNull.Value)
                                sb.Append("'").Append(r["de_ind_imp_3"].ToString()).Append("',");
                            else
                                sb.Append("null,");

                            if (r["im_base_imp1"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["im_base_imp1"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["im_base_imp2"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["im_base_imp2"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["im_base_imp3"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["im_base_imp3"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["im_costes_atr"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["im_costes_atr"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["im_impuesto_1"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["im_impuesto_1"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["im_impuesto_2"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["im_impuesto_2"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["im_impuesto_3"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["im_impuesto_3"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["nm_total_igic"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["nm_total_igic"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["nm_total_ipsi"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["nm_total_ipsi"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["nm_total_iva"] != System.DBNull.Value)
                                sb.Append(Convert.ToDouble(r["nm_total_iva"]).ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");

                            if (r["fec_act"] != System.DBNull.Value)
                                sb.Append("'").Append(Convert.ToDateTime(r["fec_act"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                            else
                                sb.Append("null,");

                            sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                            sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");
                            #endregion

                            if (j == num_reg_insert)
                            {

                                hora_inicio = DateTime.Now;
                                //db_target = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD_FACTURACION);
                                command = new OdbcCommand(sb.ToString().Substring(0, sb.Length - 1), db_target.con);
                                command.CommandTimeout = Convert.ToInt32(p.GetValue("CommandTimeout"));
                                command.ExecuteNonQuery();
                                //db_target.CloseConnection();
                                hora_fin = DateTime.Now;

                                diferencia_horas = hora_fin.Subtract(hora_inicio);
                                segundos = Convert.ToInt32(diferencia_horas.TotalSeconds);
                                segundos_totales += segundos;
                                reg_seg = (float)(num_reg_insert / segundos);
                                minutos_restantes = Convert.ToInt32((1 / reg_seg) * (totalRegistros - k) / 60);

                                Console.CursorLeft = 0;
                                Console.Write(new string(' ', Console.WindowWidth - 15));
                                Console.CursorLeft = 0;
                                Console.Write("Anexamos " + String.Format("{0:N0}", k) + "/" + String.Format("{0:N0}", totalRegistros) + " [Vel.: " + String.Format("{0:N0}", reg_seg) +
                                 " reg/s ==> Minutos restantes: " + string.Format("{0:N0}", minutos_restantes) + " | Minutos totales: " + string.Format("{0:N0}", (segundos_totales / 60)) + "]");

                                firstOnly = true;
                                sb = null;
                                sb = new StringBuilder();
                                j = 0;
                            }

                        }

                        if (j > 0)
                        {

                            hora_inicio = DateTime.Now;
                            //db_target = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD_FACTURACION);
                            command = new OdbcCommand(sb.ToString().Substring(0, sb.Length - 1), db_target.con);
                            command.ExecuteNonQuery();
                            //db_target.CloseConnection();
                            hora_fin = DateTime.Now;

                            diferencia_horas = hora_fin.Subtract(hora_inicio);
                            segundos = Convert.ToInt32(diferencia_horas.TotalSeconds);
                            reg_seg = (float)(num_reg_insert / segundos);
                            minutos_restantes = Convert.ToInt32((1 / reg_seg) * (totalRegistros - k) / 60);

                            Console.CursorLeft = 0;
                            Console.Write(new string(' ', Console.WindowWidth - 15));
                            Console.CursorLeft = 0;
                            Console.Write("Anexamos " + String.Format("{0:N0}", k) + "/" + String.Format("{0:N0}", totalRegistros) + " [Vel.: " + String.Format("{0:N0}", reg_seg) +
                                 " reg/s ==> Minutos restantes: " + string.Format("{0:N0}", minutos_restantes) + " | Minutos totales: " + string.Format("{0:N0}", (segundos_totales / 60)) + "]");

                            firstOnly = true;
                            sb = null;
                            sb = new StringBuilder();
                            j = 0;
                        }

                        ficheroLogB2B.Add("Fin anexión datos. Añadidos " + String.Format("{0:N0}", k) + " registros de Facturas B2B - DI de un total de " + String.Format("{0:N0}", totalRegistros) + " | Tiempo total: " + string.Format("{0:N0}", (segundos_totales / 60)) + " minutos.");

                        ss_pp.Update_Fecha_Fin("Facturación", "Copiar Facturas B2B", "Copiar Facturas B2B - DI");
                        #endregion
                    }
                    else
                    {
                        ss_pp.Update_Fecha_Fin("Facturación", "Copiar Facturas B2B", "Copiar Facturas B2B - DI");
                        ss_pp.Update_Comentario("Facturación", "Copiar Facturas B2B", "Copiar Facturas B2B - DI", "No se recuperaron registros");
                    }

                    db_source.CloseConnection();
                    db_target.CloseConnection();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Thread.Sleep(3000);
                    ficheroLogB2B.Add("Se produjo un error, se ha podido anexar " + String.Format("{ 0:N0}", k) + " registros de Facturas B2B - DI de un total de " + String.Format("{ 0:N0}", totalRegistros) + " | Tiempo total: " + string.Format("{ 0:N0}", (segundos_totales / 60)) + " minutos.");
                    ficheroLogB2B.AddError(ex.Message);
                    ss_pp.Update_Fecha_Fin("Facturación", "Copiar Facturas B2B", "Copiar Facturas B2B - DI");
                    ss_pp.Update_Comentario("Facturación", "Copiar Facturas B2B", "Copiar Facturas B2B - DI", "Error en CopiarFacturasB2B_t_ed_h_sap_facts()");
                }
            }
        }
        public void CopiarFacturasB2B_t_ed_h_sap_facts_conceptos()
        {
         
        }
        public void CopiarFacturasB2B_t_ed_h_sap_dc()
        {

        }
        public string Consulta_FacturasB2B_t_ed_h_sap_facts()
        {
            string strSql = "";
            
            strSql = "SELECT f.cl_empr, f.cd_mes, f.id_fact, f.cd_di, f.fh_fact, f.fh_ini_fact, f.fh_fin_fact, f.cd_tp_fact, f.cd_est_fact, f.cd_tp_cli, f.nm_ener_consmda, f.nm_ener_factda,"
                + " f.im_factdo_sin_iva, f.im_factdo_con_iva, f.cl_cli, f.id_fact_anulada, f.cd_di_sustyente, f.cd_di_sustituida, f.cd_di_anuladora,"
                + " f.de_tp_fact, f.de_estado_fact, f.cd_cuenta_contr, f.cd_ind_imp_1, f.de_ind_imp_1, f.cd_ind_imp_2, f.de_ind_imp_2, f.cd_ind_imp_3, f.de_ind_imp_3, f.im_base_imp1,"
                + " f.im_base_imp2, f.im_base_imp3, f.im_costes_atr, f.im_impuesto_1, f.im_impuesto_2, f.im_impuesto_3, f.nm_total_igic, f.nm_total_ipsi, f.nm_total_iva, f.fec_act"
                + " FROM ed_owner.t_ed_h_sap_facts f where"
                 + " f.fh_fact >= '" + p.GetValue("fecha_factura_desde") + "' AND"
                + " f.fh_fact <= '" + p.GetValue("fecha_factura_hasta") + "' AND"
                + " f.de_marca_back = '" + p.GetValue("de_marca_back") + "' AND"
                + " f.id_fact is not null AND f.id_fact <> '' ";

            return strSql;
        }
        public string Consulta_Delete_FacturasB2B_t_ed_h_sap_facts()
        {
            string strSql = "";

            strSql = "DELETE FROM facturacionb2b_owner.t_ed_h_sap_facts";

            return strSql;
        }
        public string Total_Consulta_FacturasB2B_t_ed_h_sap_facts()
        {
            string strSql = "";

            strSql = "SELECT COUNT(*) as total"
                + " FROM ed_owner.t_ed_h_sap_facts f where"
                 + " f.fh_fact >= '" + p.GetValue("fecha_factura_desde") + "' AND"
                + " f.fh_fact <= '" + p.GetValue("fecha_factura_hasta") + "' AND"
                + " f.de_marca_back = '" + p.GetValue("de_marca_back") + "' AND"
                + " f.id_fact is not null AND f.id_fact <> '' ";

            return strSql;
        }
        public string Consulta_FacturasB2B_t_ed_h_sap_facts_conceptos()
        {
            string strSql = "";

            strSql = "SELECT cd_di ";
            for (int i = 1; i <= 50; i++)
                strSql += " ,cd_concepto_" + i;
            for (int i = 1; i <= 50; i++)
                strSql += " ,de_concepto_" + i;
            for (int i = 1; i <= 50; i++)
                strSql += " ,im_concepto_" + i;
            strSql += " FROM ed_owner.t_ed_h_sap_facts where"
                + " fh_fact >= '" + p.GetValue("fecha_factura_desde") + "' AND"
                + " fh_fact <= '" + p.GetValue("fecha_factura_hasta") + "' AND"
                + " de_marca_back = '" + p.GetValue("de_marca_back")  + "' AND"
                + " id_fact is not null AND id_fact <> '' ";
                
            return strSql;
        }
        public string Consulta_Delete_FacturasB2B_t_ed_h_sap_facts_conceptos()
        {
            string strSql = "";

            strSql = "DELETE FROM facturacionb2b_owner.t_ed_h_sap_facts_conceptos";

            return strSql;
        }
        public string Total_Consulta_FacturasB2B_t_ed_h_sap_facts_conceptos()
        {
            string strSql = "";

            strSql = "SELECT COUNT(*) as total";
            strSql += " FROM ed_owner.t_ed_h_sap_facts where"
                + " fh_fact >= '" + p.GetValue("fecha_factura_desde") + "' AND"
                + " fh_fact <= '" + p.GetValue("fecha_factura_hasta") + "' AND"
                + " de_marca_back = '" + p.GetValue("de_marca_back") + "' AND"
                + " id_fact is not null AND id_fact <> '' ";

            return strSql;
        }
        public string Consulta_FacturasB2B_t_ed_h_sap_dc()
        {
            string strSql = "";

            strSql = "SELECT dc.cd_dc, dc.cd_di, dc.cd_cups_ext, dc.cd_cups_gas_ext, dc.fec_act"
                + " FROM ed_owner.t_ed_h_sap_facts f"
                + " inner join ed_owner.t_ed_h_sap_dc dc on"
                + " dc.cd_di = f.cd_di where"
                + " f.fh_fact >= '" + p.GetValue("fecha_factura_desde") + "' AND"
                + " f.fh_fact <= '" + p.GetValue("fecha_factura_hasta") + "' AND"
                + " f.de_marca_back = '" + p.GetValue("de_marca_back") + "' AND"
                + " f.id_fact is not null AND f.id_fact <> '' ";

            return strSql;
        }
        public string Consulta_Delete_FacturasB2B_t_ed_h_sap_dc()
        {
            string strSql = "";

            strSql = "DELETE FROM facturacionb2b_owner.t_ed_h_sap_dc";

            return strSql;
        }
        public string Total_Consulta_FacturasB2B_t_ed_h_sap_dc()
        {
            string strSql = "";

            strSql = "SELECT COUNT(*) as total"
                + " FROM ed_owner.t_ed_h_sap_facts f"
                + " inner join ed_owner.t_ed_h_sap_dc dc on"
                + " dc.cd_di = f.cd_di where"
                + " f.fh_fact >= '" + p.GetValue("fecha_factura_desde") + "' AND"
                + " f.fh_fact <= '" + p.GetValue("fecha_factura_hasta") + "' AND"
                + " f.de_marca_back = '" + p.GetValue("de_marca_back") + "' AND"
                + " f.id_fact is not null AND f.id_fact <> '' ";

            return strSql;
        }

    }
}
