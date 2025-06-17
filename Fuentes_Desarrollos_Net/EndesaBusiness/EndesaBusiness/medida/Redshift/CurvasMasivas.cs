using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using EndesaEntity;
using EndesaEntity.cnmc.gas.V25_2019_12_17;
using EndesaEntity.medida;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.medida.Redshift
{
    public class CurvasMasivas
    {
        utilidades.Fechas fechas;
        logs.Log ficheroLog;
        public Dictionary<string, List<CurvaCuartoHoraria>> dic_cc { get; set; }
        public Dictionary<string, List<string>> dic_peticiones;

        EndesaBusiness.medida.Redshift.Estados_Curvas estados_curvas;

        utilidades.Seguimiento_Procesos ss_pp;
        public CurvasMasivas()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_CurvasMasivas");
            fechas = new utilidades.Fechas();
            dic_cc = new Dictionary<string, List<CurvaCuartoHoraria>>();
            estados_curvas = new Redshift.Estados_Curvas();
            ss_pp = new Seguimiento_Procesos();
        }


        public void LeePeticionMasiva()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string fechas = "";
            bool hayError = false;
            DateTime fechaDesde = new DateTime();
            DateTime fechaHasta = new DateTime();
            int year = 0;
            int month = 0;
            int day = 0;

            ss_pp.Update_Fecha_Inicio("Medida", "Curvas Masivas BI", "Curvas Masivas BI");

            dic_peticiones = new Dictionary<string, List<string>>();

            //***********************************************************
            // Hacemos listas de CUPS para cada periodo de fechas iguales
            //***********************************************************

            strSql = "SELECT cups, fecha_desde, fecha_hasta FROM med.bi_peticiones_extraccion";
            db = new MySQLDB(MySQLDB.Esquemas.GBL);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                fechas = Convert.ToDateTime(r["fecha_desde"]).ToString("yyyyMMdd") + "_"
                    + Convert.ToDateTime(r["fecha_hasta"]).ToString("yyyyMMdd");

                List<string> o;
                if (!dic_peticiones.TryGetValue(fechas, out o))
                {
                    o = new List<string>();
                    o.Add(r["cups"].ToString());
                    dic_peticiones.Add(fechas, o);
                }else
                    o.Add(r["cups"].ToString());
            }
            db.CloseConnection();

            foreach(KeyValuePair<string, List<string>> p in dic_peticiones)
            {
                year = Convert.ToInt32(p.Key.Substring(0,4));
                month = Convert.ToInt32(p.Key.Substring(4, 2));
                day = Convert.ToInt32(p.Key.Substring(6, 2));
                fechaDesde = new DateTime(year, month, day);
                year = Convert.ToInt32(p.Key.Substring(9, 4));
                month = Convert.ToInt32(p.Key.Substring(13, 2));
                day = Convert.ToInt32(p.Key.Substring(15, 2));
                fechaHasta = new DateTime(year, month, day);

                hayError = GetRCurva(estados_curvas, p.Value, fechaDesde, fechaHasta, estados_curvas.estados_registrados);
                hayError = GetRCurva(estados_curvas, p.Value, fechaDesde, fechaHasta, estados_curvas.estados_facturados);
                hayError = GetCurva(estados_curvas, p.Value, fechaDesde, fechaHasta, estados_curvas.estados_registrados);
                hayError = GetCurva(estados_curvas, p.Value, fechaDesde, fechaHasta, estados_curvas.estados_facturados);

                ss_pp.Update_Fecha_Fin("Medida", "Curvas Masivas BI", "Curvas Masivas BI");
            }
            

        }


        public bool GetCurva(Estados_Curvas estado_curvas, List<string> lista, DateTime fd, DateTime fh, List<string> lista_estados)
        {
            string strSql = "";
            StringBuilder sb = new StringBuilder();
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;


            MySQLDB dbm;
            MySqlCommand commandm;
            

            int numeroPeriodos = 0;
            int year = 0;
            int month = 0;
            int day = 0;
            int x = 0;
            int i = 0;
            
            DateTime fechaHora = new DateTime();
            bool firstOnly = true;

            try
            {

                #region Consulta
                Console.WriteLine("Realizando consulta de " + lista.Count()
                    + " cups y fechas " + fd.ToString("dd/MM/yyyy") + " a " + fh.ToString("dd/MM/yyyy"));

                sb.Append("SELECT cd_empr_titular, cd_particion, cd_finca, cd_pto_servicio, cd_punto_med, fh_lect_registro, cd_sec_resumen,");
                sb.Append("cd_distrib, cd_cups, cd_cups_ext, cd_cups_ext_20, cd_estado_curva,");

                for (int h = 1; h <= 25; h++)
                {
                    for (int c = 1; c <= 4; c++)
                        sb.Append("nm_pot_ac_h" + h + "_cuad" + c + ",");

                    sb.Append("cd_fuente_cuarth_ac_h" + h + "_cuad1,");

                }

                for (int h = 1; h <= 25; h++)
                {
                    sb.Append("nm_ener_ac_h" + h + ",");
                    sb.Append("cd_estado_med_ac_h" + h + ",");
                    sb.Append("cd_fuente_hor_ac_h" + h + ",");
                    sb.Append("cd_ind_hor_inv_ac_h" + h + ",");
                }


                for (int h = 1; h <= 25; h++)
                {
                    for (int c = 1; c <= 4; c++)
                        sb.Append("nm_pot_r1_h" + h + "_cuad" + c + ",");

                    sb.Append("cd_fuente_cuarth_r1_h" + h + "_cuad1,");

                }


                for (int h = 1; h <= 25; h++)
                {
                    sb.Append("nm_ener_r1_h" + h + ",");
                    sb.Append("cd_estado_med_r1_h" + h + ",");
                    sb.Append("cd_fuente_hor_r1_h" + h + ",");
                    sb.Append("cd_ind_hor_inv_r1_h" + h + ",");
                }

                sb.Append("fh_act_regist_metra, cd_tp_curva, lg_perdidas, de_estado_curva, id_huella_periodo, cd_est_facturacion_sap,");
                sb.Append("de_est_facturacion_sap, id_dc, id_fact, fh_fact,");

                //for (int h = 1; h <= 25; h++)                
                //    for (int c = 1; c <= 4; c++)
                //        sb.Append("nm_pot_r2_h" + h + "_cuad" + c + ",");

                //for (int h = 1; h <= 25; h++)                    
                //    sb.Append("nm_ener_r2_h" + h + ",");

                //for (int h = 1; h <= 25; h++)
                //    for (int c = 1; c <= 4; c++)
                //        sb.Append("nm_pot_r3_h" + h + "_cuad" + c + ",");

                //for (int h = 1; h <= 25; h++)
                //    sb.Append("nm_ener_r3_h" + h + ",");

                for (int h = 1; h <= 25; h++)
                    for (int c = 1; c <= 4; c++)
                        sb.Append("nm_pot_r4_h" + h + "_cuad" + c + ",");


                for (int h = 1; h <= 25; h++)
                    sb.Append("nm_ener_r4_h" + h + ",");


                for (int h = 1; h <= 25; h++)
                    for (int c = 1; c <= 4; c++)
                        sb.Append("nm_pot_as_h" + h + "_cuad" + c + ",");

                for (int h = 1; h <= 25; h++)
                    sb.Append("nm_ener_as_h" + h + ",");


                sb.Append(" cod_carga, fh_recepcion");
                sb.Append(" FROM METRA_OWNER.T_ED_H_CURVAS WHERE");
                sb.Append(" cd_cups_ext_20 in (");;

                for (int y = 0; y < lista.Count; y++)
                {
                    if (firstOnly)
                    {
                        sb.Append("'").Append(lista[y]).Append("'");
                        firstOnly = false;
                    }
                    else
                        sb.Append(",'").Append(lista[y]).Append("'");

                }

                sb.Append(") AND");
                sb.Append(" (FH_LECT_REGISTRO >= ").Append(fd.ToString("yyyyMMdd"));
                sb.Append(" AND FH_LECT_REGISTRO <= ").Append(fh.ToString("yyyyMMdd")).Append(")");
                sb.Append(" and CD_ESTADO_CURVA in ('").Append(lista_estados[0]).Append("'");

                for (int w = 1; w < lista_estados.Count; w++)
                    sb.Append(",'").Append(lista_estados[w]).Append("'");

                sb.Append(") ORDER BY CD_CUPS_EXT, FH_LECT_REGISTRO, CD_SEC_RESUMEN DESC;");


                ficheroLog.Add(sb.ToString());
                #endregion

                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();


                firstOnly = true;
                sb = null;
                sb = new StringBuilder();
                while (r.Read())
                {

                    x++;
                    i++;

                    if (firstOnly)
                    {
                        sb.Append("REPLACE INTO t_ed_h_curvas ");
                        sb.Append("(cd_empr_titular, cd_particion, cd_finca, cd_pto_servicio, cd_punto_med, fh_lect_registro, cd_sec_resumen,");
                        sb.Append("cd_distrib, cd_cups, cd_cups_ext, cd_cups_ext_20, cd_estado_curva,");

                        for (int h = 1; h <= 25; h++)
                        {
                            for (int c = 1; c <= 4; c++)
                                sb.Append("nm_pot_ac_h" + h + "_cuad" + c + ",");

                            sb.Append("cd_fuente_cuarth_ac_h" + h + "_cuad1,");

                        }

                        for (int h = 1; h <= 25; h++)
                        {
                            sb.Append("nm_ener_ac_h" + h + ",");
                            sb.Append("cd_estado_med_ac_h" + h + ",");
                            sb.Append("cd_fuente_hor_ac_h" + h + ",");
                            sb.Append("cd_ind_hor_inv_ac_h" + h + ",");
                        }


                        for (int h = 1; h <= 25; h++)
                        {
                            for (int c = 1; c <= 4; c++)
                                sb.Append("nm_pot_r1_h" + h + "_cuad" + c + ",");

                            sb.Append("cd_fuente_cuarth_r1_h" + h + "_cuad1,");

                        }


                        for (int h = 1; h <= 25; h++)
                        {
                            sb.Append("nm_ener_r1_h" + h + ",");
                            sb.Append("cd_estado_med_r1_h" + h + ",");
                            sb.Append("cd_fuente_hor_r1_h" + h + ",");
                            sb.Append("cd_ind_hor_inv_r1_h" + h + ",");
                        }

                        sb.Append("fh_act_regist_metra, cd_tp_curva, lg_perdidas, de_estado_curva, id_huella_periodo, cd_est_facturacion_sap,");
                        sb.Append("de_est_facturacion_sap, id_dc, id_fact, fh_fact,");

                        //for (int h = 1; h <= 25; h++)
                        //    for (int c = 1; c <= 4; c++)
                        //        sb.Append("nm_pot_r2_h" + h + "_cuad" + c + ",");

                        //for (int h = 1; h <= 25; h++)
                        //    sb.Append("nm_ener_r2_h" + h + ",");

                        //for (int h = 1; h <= 25; h++)
                        //    for (int c = 1; c <= 4; c++)
                        //        sb.Append("nm_pot_r3_h" + h + "_cuad" + c + ",");

                        //for (int h = 1; h <= 25; h++)
                        //    sb.Append("nm_ener_r3_h" + h + ",");

                        for (int h = 1; h <= 25; h++)
                            for (int c = 1; c <= 4; c++)
                                sb.Append("nm_pot_r4_h" + h + "_cuad" + c + ",");


                        for (int h = 1; h <= 25; h++)
                            sb.Append("nm_ener_r4_h" + h + ",");


                        for (int h = 1; h <= 25; h++)
                            for (int c = 1; c <= 4; c++)
                                sb.Append("nm_pot_as_h" + h + "_cuad" + c + ",");

                        for (int h = 1; h <= 25; h++)
                            sb.Append("nm_ener_as_h" + h + ",");


                        sb.Append(" cod_carga, fh_recepcion, created_by, created_date) values ");
                        firstOnly = false;
                    }


                    sb.Append("('").Append(r["cd_empr_titular"].ToString()).Append("',");
                    sb.Append(r["cd_particion"].ToString()).Append(",");
                    sb.Append(r["cd_finca"].ToString()).Append(",");
                    sb.Append(r["cd_pto_servicio"].ToString()).Append(",");
                    sb.Append("'").Append(r["cd_punto_med"].ToString()).Append("',");
                    sb.Append(r["fh_lect_registro"].ToString()).Append(",");
                    sb.Append(r["cd_sec_resumen"].ToString()).Append(",");
                    sb.Append("'").Append(r["cd_distrib"].ToString()).Append("',");
                    sb.Append("'").Append(r["cd_cups"].ToString()).Append("',");
                    sb.Append("'").Append(r["cd_cups_ext"].ToString()).Append("',");
                    sb.Append("'").Append(r["cd_cups_ext_20"].ToString()).Append("',");
                    sb.Append("'").Append(r["cd_estado_curva"].ToString().Trim()).Append("',");


                    for (int h = 1; h <= 25; h++)
                    {
                        for (int c = 1; c <= 4; c++)
                        {
                            if (r["nm_pot_ac_h" + h + "_cuad" + c] != System.DBNull.Value)
                                sb.Append(r["nm_pot_ac_h" + h + "_cuad" + c].ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");
                        }

                        if (r["cd_fuente_cuarth_ac_h" + h + "_cuad1"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_fuente_cuarth_ac_h" + h + "_cuad1"].ToString().Trim()).Append("',");
                        else
                            sb.Append("null,");

                    }


                    for (int h = 1; h <= 25; h++)
                    {
                        if (r["nm_ener_ac_h" + h] != System.DBNull.Value)
                            sb.Append(r["nm_ener_ac_h" + h].ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["cd_estado_med_ac_h" + h] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_estado_med_ac_h" + h].ToString().Trim()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_fuente_hor_ac_h" + h] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_fuente_hor_ac_h" + h].ToString().Trim()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_ind_hor_inv_ac_h" + h] != System.DBNull.Value)
                            sb.Append(r["cd_ind_hor_inv_ac_h" + h].ToString()).Append(",");
                        else
                            sb.Append("null,");
                    }


                    for (int h = 1; h <= 25; h++)
                    {
                        for (int c = 1; c <= 4; c++)
                        {
                            if (r["nm_pot_r1_h" + h + "_cuad" + c] != System.DBNull.Value)
                                sb.Append(r["nm_pot_r1_h" + h + "_cuad" + c].ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");
                        }

                        if (r["cd_fuente_cuarth_r1_h" + h + "_cuad1"] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_fuente_cuarth_r1_h" + h + "_cuad1"].ToString().Trim()).Append("',");
                        else
                            sb.Append("null,");

                    }

                    for (int h = 1; h <= 25; h++)
                    {
                        if (r["nm_ener_r1_h" + h] != System.DBNull.Value)
                            sb.Append(r["nm_ener_r1_h" + h].ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["cd_estado_med_r1_h" + h] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_estado_med_r1_h" + h].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_fuente_hor_r1_h" + h] != System.DBNull.Value)
                            sb.Append("'").Append(r["cd_fuente_hor_r1_h" + h].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        if (r["cd_ind_hor_inv_r1_h" + h] != System.DBNull.Value)
                            sb.Append(r["cd_ind_hor_inv_r1_h" + h].ToString()).Append(",");
                        else
                            sb.Append("null,");
                    }


                    if (r["fh_act_regist_metra"] != System.DBNull.Value)
                        sb.Append(r["fh_act_regist_metra"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_curva"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_curva"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_perdidas"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_perdidas"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_estado_curva"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_estado_curva"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_huella_periodo"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_huella_periodo"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_est_facturacion_sap"] != System.DBNull.Value)
                        sb.Append(r["cd_est_facturacion_sap"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["de_est_facturacion_sap"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_est_facturacion_sap"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_dc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_dc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_fact"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");



                    //for (int h = 1; h <= 25; h++)
                    //    for (int c = 1; c <= 4; c++)
                    //    {
                    //        if (r["nm_pot_r2_h" + h + "_cuad" + c] != System.DBNull.Value)
                    //            sb.Append(r["nm_pot_r2_h" + h + "_cuad" + c].ToString().Replace(",", ".")).Append(",");
                    //        else
                    //            sb.Append("null,");
                    //    }
                            

                    //for (int h = 1; h <= 25; h++)
                    //{
                    //    if (r["nm_ener_r2_h" + h] != System.DBNull.Value)
                    //        sb.Append(r["nm_ener_r2_h" + h].ToString().Replace(",", ".")).Append(",");
                    //    else
                    //        sb.Append("null,");
                    //}
                        

                    //for (int h = 1; h <= 25; h++)
                    //    for (int c = 1; c <= 4; c++)
                    //    {
                    //        if (r["nm_pot_r3_h" + h + "_cuad" + c] != System.DBNull.Value)
                    //            sb.Append(r["nm_pot_r3_h" + h + "_cuad" + c].ToString().Replace(",", ".")).Append(",");
                    //        else
                    //            sb.Append("null,");
                    //    }
                            

                    //for (int h = 1; h <= 25; h++)
                    //{
                    //    if (r["nm_ener_r3_h" + h] != System.DBNull.Value)
                    //        sb.Append(r["nm_ener_r3_h" + h].ToString().Replace(",", ".")).Append(",");
                    //    else
                    //        sb.Append("null,");
                    //}
                        

                    for (int h = 1; h <= 25; h++)
                        for (int c = 1; c <= 4; c++)
                        {
                            if (r["nm_pot_r4_h" + h + "_cuad" + c] != System.DBNull.Value)
                                sb.Append(r["nm_pot_r4_h" + h + "_cuad" + c].ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");
                        }
                            


                    for (int h = 1; h <= 25; h++)
                    {
                        if (r["nm_ener_r4_h" + h] != System.DBNull.Value)
                            sb.Append(r["nm_ener_r4_h" + h].ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }                       


                    for (int h = 1; h <= 25; h++)
                        for (int c = 1; c <= 4; c++)
                        {
                            if (r["nm_pot_as_h" + h + "_cuad" + c] != System.DBNull.Value)
                                sb.Append(r["nm_pot_as_h" + h + "_cuad" + c].ToString().Replace(",", ".")).Append(",");
                            else
                                sb.Append("null,");
                        }
                            

                    for (int h = 1; h <= 25; h++)
                    {
                        if (r["nm_ener_as_h" + h] != System.DBNull.Value)
                            sb.Append(r["nm_ener_as_h" + h].ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");
                    }


                    if (r["cod_carga"] != System.DBNull.Value)
                        sb.Append(r["cod_carga"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_recepcion"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_recepcion"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                    if (x == 250)
                    {
                        Console.WriteLine("Guardando " + i + " registros...");
                        firstOnly = true;
                        dbm = new MySQLDB(MySQLDB.Esquemas.MED);
                        commandm = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbm.con);
                        commandm.ExecuteNonQuery();
                        dbm.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        x = 0;
                    }                  


                }

                db.CloseConnection();



                if (x > 0)
                {
                    Console.WriteLine("Guardando " + i + " registros...");
                    firstOnly = true;
                    dbm = new MySQLDB(MySQLDB.Esquemas.MED);
                    commandm = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbm.con);
                    commandm.ExecuteNonQuery();
                    dbm.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    x = 0;
                }

                return false;

            }
            catch (Exception e)
            {

                ss_pp.Update_Comentario("Medida", "Curvas Masivas BI", "Curvas Masivas BI", e.Message);
                ficheroLog.AddError("CurvasRedShift.GetCurva: " + e.Message);
                return true;
            }

        }

       

        public bool GetRCurva(Estados_Curvas estado_curvas, List<string> lista, DateTime fd, DateTime fh, List<string> lista_estados)
        {
            string strSql = "";
            StringBuilder sb = new StringBuilder();
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;


            MySQLDB dbm;
            MySqlCommand commandm;


            int numeroPeriodos = 0;
            int year = 0;
            int month = 0;
            int day = 0;
            int x = 0;
            int i = 0;

            DateTime fechaHora = new DateTime();
            bool firstOnly = true;

            try
            {

                #region Consulta
                Console.WriteLine("Realizando consulta de " + lista.Count()
                    + " cups y fechas " + fd.ToString("dd/MM/yyyy") + " a " + fh.ToString("dd/MM/yyyy"));

                sb.Append("SELECT cd_empr_titular, cd_particion, cd_finca, cd_pto_servicio, cd_punto_med, fh_fact_desde, fh_fact_hasta,");
                sb.Append(" cd_sec_resumen, cd_distrib, cd_cups, cd_cups_ext, cd_cups_ext_20, fh_lect_registro, nm_tareas_ltp, nm_curvas_recibidas,");
                sb.Append(" cd_estado_resumen, de_estado_resumen, cd_tp_fuente_horaria, de_tp_fuente_horaria, cd_tp_fuente_cuartoh, de_tp_fuente_cuartoh,");
                sb.Append(" cd_est_curva_valid, nm_cons_total_act, nm_cons_total_reac, cd_ind_objeto, tx_secuenciales_cch, cd_ind_nivel_hist, fh_fact,");
                sb.Append(" cd_usuario_sce, cd_program_sce, fh_ult_modif_sce, fh_act_regist_metra, id_huella_periodo, id_huella_ol, nm_cons_total_perdidas_act,");
                sb.Append(" nm_cons_total_perdidas_reac, nm_activa_saliente, nm_capacitiva, nm_pot_trafo, nm_por_perd, fh_generacion, lg_periodo_activo, id_prelacion,");
                sb.Append(" de_prelacion, fh_generacion_segme, tp_origen_segm, de_origen_segm, de_tipo_mov, nm_ritmo_fact, fh_lim_recep, fh_envio_ack, de_estado_ack,");
                sb.Append(" de_error_ack, cd_est_facturacion_sap, de_est_facturacion_sap, id_dc, id_fact, im_fact, cd_est_fact, de_est_fact, id_fact_sust, de_anom_sap,");
                sb.Append(" de_mot_desec, de_text_revision, nm_tam, fh_corte_fuera_plz, lg_carta_anexa, de_cod_exp_ayf, de_estado_sol_ayf, de_proc_mod, de_comentario,");
                sb.Append(" lg_carga_manual, lg_modo_proc, fh_act, cod_carga, id_periodo_kee, fh_estado, id_crto_ext, nm_sec_crto");
                sb.Append(" FROM metra_owner.t_ed_h_rcurvas WHERE");                
                sb.Append(" cd_cups_ext_20 in ("); ;

                for (int y = 0; y < lista.Count; y++)
                {
                    if (firstOnly)
                    {
                        sb.Append("'").Append(lista[y]).Append("'");
                        firstOnly = false;
                    }
                    else
                        sb.Append(",'").Append(lista[y]).Append("'");

                }

                sb.Append(") AND");
                sb.Append(" (fh_fact_desde >= ").Append(fd.ToString("yyyyMMdd"));
                sb.Append(" AND fh_fact_hasta <= ").Append(fh.ToString("yyyyMMdd")).Append(")");
                sb.Append(" and cd_estado_resumen in ('").Append(lista_estados[0]).Append("'");

                for (int w = 1; w < lista_estados.Count; w++)
                    sb.Append(",'").Append(lista_estados[w]).Append("'");

                sb.Append(") ORDER BY cd_cups_ext, fh_lect_registro, cd_sec_resumen DESC;");


                ficheroLog.Add(sb.ToString());
                #endregion

                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();


                firstOnly = true;
                sb = null;
                sb = new StringBuilder();
                while (r.Read())
                {

                    x++;
                    i++;

                    if (firstOnly)
                    {
                        sb.Append("REPLACE INTO t_ed_h_rcurvas ");
                        sb.Append("(cd_empr_titular, cd_particion, cd_finca, cd_pto_servicio, cd_punto_med,");
                        sb.Append("fh_fact_desde, fh_fact_hasta, cd_sec_resumen, cd_distrib, cd_cups, cd_cups_ext,");
                        sb.Append("cd_cups_ext_20, fh_lect_registro, nm_tareas_ltp, nm_curvas_recibidas, cd_estado_resumen,");
                        sb.Append("de_estado_resumen, cd_tp_fuente_horaria, de_tp_fuente_horaria, cd_tp_fuente_cuartoh,");
                        sb.Append("de_tp_fuente_cuartoh, cd_est_curva_valid, nm_cons_total_act, nm_cons_total_reac,");
                        sb.Append("cd_ind_objeto, tx_secuenciales_cch, cd_ind_nivel_hist, fh_fact, cd_usuario_sce,");
                        sb.Append("cd_program_sce, fh_ult_modif_sce, fh_act_regist_metra, id_huella_periodo, id_huella_ol,");
                        sb.Append("nm_cons_total_perdidas_act, nm_cons_total_perdidas_reac, nm_activa_saliente, nm_capacitiva,");
                        sb.Append("nm_pot_trafo, nm_por_perd, fh_generacion, lg_periodo_activo, id_prelacion, de_prelacion,");
                        sb.Append("fh_generacion_segme, tp_origen_segm, de_origen_segm, de_tipo_mov, nm_ritmo_fact,");
                        sb.Append("fh_lim_recep, fh_envio_ack, de_estado_ack, de_error_ack, cd_est_facturacion_sap,");
                        sb.Append("de_est_facturacion_sap, id_dc, id_fact, im_fact, cd_est_fact, de_est_fact, id_fact_sust,");
                        sb.Append("de_anom_sap, de_mot_desec, de_text_revision, nm_tam, fh_corte_fuera_plz, lg_carta_anexa,");
                        sb.Append("de_cod_exp_ayf, de_estado_sol_ayf, de_proc_mod, de_comentario, lg_carga_manual, lg_modo_proc,");
                        sb.Append("fh_act, cod_carga, id_periodo_kee, fh_estado, id_crto_ext, nm_sec_crto) values ");
                            
                        firstOnly = false;
                    }


                    sb.Append("('").Append(r["cd_empr_titular"].ToString()).Append("',");
                    sb.Append(r["cd_particion"].ToString()).Append(",");
                    sb.Append(r["cd_finca"].ToString()).Append(",");
                    sb.Append(r["cd_pto_servicio"].ToString()).Append(",");
                    sb.Append("'").Append(r["cd_punto_med"].ToString()).Append("',");
                    sb.Append(r["fh_fact_desde"].ToString()).Append(",");
                    sb.Append(r["fh_fact_hasta"].ToString()).Append(",");
                    sb.Append(r["cd_sec_resumen"].ToString()).Append(",");
                    sb.Append("'").Append(r["cd_distrib"].ToString()).Append("',");
                    sb.Append("'").Append(r["cd_cups"].ToString()).Append("',");
                    sb.Append("'").Append(r["cd_cups_ext"].ToString()).Append("',");
                    sb.Append("'").Append(r["cd_cups_ext_20"].ToString()).Append("',");
                    sb.Append(r["fh_lect_registro"].ToString()).Append(",");

                    if (r["nm_tareas_ltp"] != System.DBNull.Value)
                        sb.Append(r["nm_tareas_ltp"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_curvas_recibidas"] != System.DBNull.Value)
                        sb.Append(r["nm_curvas_recibidas"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_estado_resumen"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_estado_resumen"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_estado_resumen"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_estado_resumen"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_horaria"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_horaria"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_fuente_horaria"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_fuente_horaria"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_tp_fuente_cuartoh"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_tp_fuente_cuartoh"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tp_fuente_cuartoh"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tp_fuente_cuartoh"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_est_curva_valid"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_est_curva_valid"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_cons_total_act"] != System.DBNull.Value)
                        sb.Append(r["nm_cons_total_act"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_cons_total_reac"] != System.DBNull.Value)
                        sb.Append(r["nm_cons_total_reac"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_ind_objeto"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_ind_objeto"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["tx_secuenciales_cch"] != System.DBNull.Value)
                        sb.Append("'").Append(r["tx_secuenciales_cch"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_ind_nivel_hist"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_ind_nivel_hist"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_fact"] != System.DBNull.Value)
                        sb.Append(r["fh_fact"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_usuario_sce"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_usuario_sce"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_program_sce"] != System.DBNull.Value)
                        sb.Append("'").Append(r["cd_program_sce"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_ult_modif_sce"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_ult_modif_sce"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_act_regist_metra"] != System.DBNull.Value)
                        sb.Append(r["fh_act_regist_metra"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["id_huella_periodo"] != System.DBNull.Value)
                        sb.Append(r["id_huella_periodo"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["id_huella_ol"] != System.DBNull.Value)
                        sb.Append(r["id_huella_ol"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_cons_total_perdidas_act"] != System.DBNull.Value)
                        sb.Append(r["nm_cons_total_perdidas_act"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_cons_total_perdidas_reac"] != System.DBNull.Value)
                        sb.Append(r["nm_cons_total_perdidas_reac"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_activa_saliente"] != System.DBNull.Value)
                        sb.Append(r["nm_activa_saliente"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_capacitiva"] != System.DBNull.Value)
                        sb.Append(r["nm_capacitiva"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_pot_trafo"] != System.DBNull.Value)
                        sb.Append(r["nm_pot_trafo"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["nm_por_perd"] != System.DBNull.Value)
                        sb.Append(r["nm_por_perd"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_generacion"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_generacion"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_periodo_activo"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_periodo_activo"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_prelacion"] != System.DBNull.Value)
                        sb.Append(r["id_prelacion"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["de_prelacion"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_prelacion"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_generacion_segme"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_generacion_segme"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["tp_origen_segm"] != System.DBNull.Value)
                        sb.Append(r["tp_origen_segm"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["de_origen_segm"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_origen_segm"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_tipo_mov"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_tipo_mov"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_ritmo_fact"] != System.DBNull.Value)
                        sb.Append(r["nm_ritmo_fact"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_lim_recep"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_lim_recep"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_envio_ack"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_envio_ack"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_estado_ack"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_estado_ack"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_error_ack"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_error_ack"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_est_facturacion_sap"] != System.DBNull.Value)
                        sb.Append(r["cd_est_facturacion_sap"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["de_est_facturacion_sap"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_est_facturacion_sap"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_dc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_dc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["im_fact"] != System.DBNull.Value)
                        sb.Append(r["im_fact"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["cd_est_fact"] != System.DBNull.Value)
                        sb.Append(r["cd_est_fact"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["de_est_fact"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_est_fact"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["id_fact_sust"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_fact_sust"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_anom_sap"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_anom_sap"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_mot_desec"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_mot_desec"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_text_revision"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_text_revision"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_tam"] != System.DBNull.Value)
                        sb.Append(r["nm_tam"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_corte_fuera_plz"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_corte_fuera_plz"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_carta_anexa"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_carta_anexa"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_cod_exp_ayf"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_cod_exp_ayf"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_estado_sol_ayf"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_estado_sol_ayf"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_proc_mod"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_proc_mod"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_comentario"] != System.DBNull.Value)
                        sb.Append("'").Append(r["de_comentario"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_carga_manual"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_carga_manual"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["lg_modo_proc"] != System.DBNull.Value)
                        sb.Append("'").Append(r["lg_modo_proc"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fh_act"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_act"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cod_carga"] != System.DBNull.Value)
                        sb.Append(r["cod_carga"].ToString()).Append(",");
                    else
                        sb.Append("null,");

                    if (r["id_periodo_kee"] != System.DBNull.Value)
                        sb.Append(r["id_periodo_kee"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (r["fh_estado"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fh_estado"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");                   

                    if (r["id_crto_ext"] != System.DBNull.Value)
                        sb.Append("'").Append(r["id_crto_ext"].ToString()).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nm_sec_crto"] != System.DBNull.Value)
                        sb.Append(r["nm_sec_crto"].ToString().Replace(",", ".")).Append("),");
                    else
                        sb.Append("null),");

                    if (x == 250)
                    {
                        Console.WriteLine("Guardando " + i + " registros...");
                        firstOnly = true;
                        dbm = new MySQLDB(MySQLDB.Esquemas.MED);
                        commandm = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbm.con);
                        commandm.ExecuteNonQuery();
                        dbm.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        x = 0;
                    }


                }

                db.CloseConnection();



                if (x > 0)
                {
                    Console.WriteLine("Guardando " + i + " registros...");
                    firstOnly = true;
                    dbm = new MySQLDB(MySQLDB.Esquemas.MED);
                    commandm = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbm.con);
                    commandm.ExecuteNonQuery();
                    dbm.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    x = 0;
                }

                return false;

            }
            catch (Exception e)
            {

                ss_pp.Update_Comentario("Medida", "Curvas Masivas BI", "Curvas Masivas BI", e.Message);
                ficheroLog.AddError("CurvasRedShift.GetCurva: " + e.Message);
                return true;
            }

        }
    }
}
