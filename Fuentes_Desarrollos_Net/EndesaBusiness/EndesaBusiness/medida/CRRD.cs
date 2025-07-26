using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Odbc;

namespace EndesaBusiness.medida
{
    public class CRRD
    {


        logs.Log ficheroLog;
        Dictionary<string, List<EndesaEntity.medida.CurvaResumen>> dic_curva_resumen;
        
        public CRRD()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_RedShist_CC");
        }        
        
        public CRRD(List<string> lista_cups20, DateTime fd, DateTime fh, string estado)
        {
            dic_curva_resumen = ConsultaCurvaResumen(lista_cups20, fd, fh, estado);
            
        }
        public CRRD(List<string> lista_cups13, string estado)
        {
            dic_curva_resumen = ConsultaCurvaResumen(lista_cups13, estado);

        }


        public void CopiaCurvasDeCargaResumen()
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;

            servidores.MySQLDB mdb;
            MySqlCommand mcommand;
            StringBuilder sb = new StringBuilder();
            int total = 0;
            int totalReg = 0;
            int registrosLeidos = 0;
            bool firstOnly = true;
            string ff = "";
            string strSql = "";

            strSql = "SELECT COUNT(*) TOTAL_REGISTROS FROM DTMCO_PRO.METRA_OWNER.T_ED_H_RCURVAS"
               + " WHERE fh_act_regist_metra >= " + UltimoDiaCurvasResumen();

            Console.WriteLine(strSql);
            ficheroLog.Add(strSql);
            db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
            command = new OdbcCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
                total = Convert.ToInt32(r["TOTAL_REGISTROS"]);
            db.CloseConnection();


            strSql = "SELECT CD_EMPR_TITULAR,CD_PARTICION,CD_FINCA,CD_PTO_SERVICIO,CD_PUNTO_MED,FH_FACT_DESDE,FH_FACT_HASTA,"
                + "CD_SEC_RESUMEN,CD_DISTRIB,CD_CUPS,CD_CUPS_EXT,CD_CUPS_EXT_20,FH_LECT_REGISTRO,NM_TAREAS_LTP,"
                + "NM_CURVAS_RECIBIDAS,CD_ESTADO_RESUMEN,DE_ESTADO_RESUMEN,CD_TP_FUENTE_HORARIA,DE_TP_FUENTE_HORARIA,"
                + "CD_TP_FUENTE_CUARTOH,DE_TP_FUENTE_CUARTOH,CD_EST_CURVA_VALID,NM_CONS_TOTAL_ACT,NM_CONS_TOTAL_REAC,"
                + "CD_IND_OBJETO,TX_SECUENCIALES_CCH,CD_IND_NIVEL_HIST,FH_FACT,CD_USUARIO_SCE,CD_PROGRAM_SCE,"
                + "FH_ULT_MODIF_SCE,FH_ACT_REGIST_METRA FROM DTMCO_PRO.METRA_OWNER.T_ED_H_RCURVAS"
                + " WHERE fh_act_regist_metra >= " + UltimoDiaCurvasResumen();
            Console.WriteLine(strSql);
            Console.WriteLine();
            ficheroLog.Add(strSql);
            db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
            command = new OdbcCommand(strSql, db.con);
            r = command.ExecuteReader();

            while (r.Read())
            {

                registrosLeidos++;
                totalReg++;
                if (firstOnly)
                {
                    sb.Append("replace into dt_t_ed_h_rcurvas (cd_empr_titular, cd_particion, cd_finca,");
                    sb.Append("cd_pto_servicio , cd_punto_med, fh_fact_desde,");
                    sb.Append("fh_fact_hasta , cd_sec_resumen, cd_distrib,");
                    sb.Append("cd_cups , cd_cups_ext, cd_cups_ext_20,");
                    sb.Append("fh_lect_registro , nm_tareas_ltp, nm_curvas_recibidas,");
                    sb.Append("cd_estado_resumen , de_estado_resumen, cd_tp_fuente_horaria,");
                    sb.Append("de_tp_fuente_horaria , cd_tp_fuente_cuartoh, de_tp_fuente_cuartoh,");
                    sb.Append("cd_est_curva_valid , nm_cons_total_act, nm_cons_total_reac,");
                    sb.Append("cd_ind_objeto , tx_secuenciales_cch, cd_ind_nivel_hist,");
                    sb.Append("fh_fact , cd_usuario_sce, cd_program_sce,");
                    sb.Append("fh_ult_modif_sce, fh_act_regist_metra) values ");
                    firstOnly = false;
                }

                if (r["CD_EMPR_TITULAR"] != System.DBNull.Value)
                    sb.Append("('").Append(r["CD_EMPR_TITULAR"].ToString()).Append("'");
                else
                    sb.Append("(null");

                if (r["CD_PARTICION"] != System.DBNull.Value)
                    sb.Append(", ").Append(r["CD_PARTICION"].ToString());
                else
                    sb.Append(",null");

                if (r["CD_FINCA"] != System.DBNull.Value)
                    sb.Append(", ").Append(r["CD_FINCA"].ToString());
                else
                    sb.Append(",null");


                if (r["CD_PTO_SERVICIO"] != System.DBNull.Value)
                    sb.Append(", ").Append(r["CD_PTO_SERVICIO"].ToString());
                else
                    sb.Append(",null");


                if (r["CD_PUNTO_MED"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["CD_PUNTO_MED"].ToString()).Append("'");
                else
                    sb.Append(",null");


                if (r["FH_FACT_DESDE"] != System.DBNull.Value)
                {
                    ff = r["FH_FACT_DESDE"].ToString();
                    sb.Append(", '").Append(ff).Append("'");
                }
                else
                    sb.Append(",null");




                if (r["FH_FACT_HASTA"] != System.DBNull.Value)
                {
                    ff = r["FH_FACT_HASTA"].ToString();
                    sb.Append(", '").Append(ff).Append("'");
                }
                else
                    sb.Append(",null");



                if (r["CD_SEC_RESUMEN"] != System.DBNull.Value)
                    sb.Append(", ").Append(r["CD_SEC_RESUMEN"].ToString());
                else
                    sb.Append(",null");

                if (r["CD_DISTRIB"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["CD_DISTRIB"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["CD_CUPS"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["CD_CUPS"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["CD_CUPS_EXT"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["CD_CUPS_EXT"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["CD_CUPS_EXT_20"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["CD_CUPS_EXT_20"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["FH_LECT_REGISTRO"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["FH_LECT_REGISTRO"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["NM_TAREAS_LTP"] != System.DBNull.Value)
                    sb.Append(", ").Append(r["NM_TAREAS_LTP"].ToString());
                else
                    sb.Append(",null");

                if (r["NM_CURVAS_RECIBIDAS"] != System.DBNull.Value)
                    sb.Append(", ").Append(r["NM_CURVAS_RECIBIDAS"].ToString());
                else
                    sb.Append(",null");

                if (r["CD_ESTADO_RESUMEN"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["CD_ESTADO_RESUMEN"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["DE_ESTADO_RESUMEN"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["DE_ESTADO_RESUMEN"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["CD_TP_FUENTE_HORARIA"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["CD_TP_FUENTE_HORARIA"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["DE_TP_FUENTE_HORARIA"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["DE_TP_FUENTE_HORARIA"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["CD_TP_FUENTE_CUARTOH"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["CD_TP_FUENTE_CUARTOH"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["DE_TP_FUENTE_CUARTOH"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["DE_TP_FUENTE_CUARTOH"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["CD_EST_CURVA_VALID"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["CD_EST_CURVA_VALID"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["NM_CONS_TOTAL_ACT"] != System.DBNull.Value)
                    sb.Append(", ").Append(r["NM_CONS_TOTAL_ACT"].ToString().Replace(",","."));
                else
                    sb.Append(",null");

                if (r["NM_CONS_TOTAL_REAC"] != System.DBNull.Value)
                    sb.Append(", ").Append(r["NM_CONS_TOTAL_REAC"].ToString().Replace(",", "."));
                else
                    sb.Append(",null");

                if (r["CD_IND_OBJETO"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["CD_IND_OBJETO"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["TX_SECUENCIALES_CCH"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["TX_SECUENCIALES_CCH"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["CD_IND_NIVEL_HIST"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["CD_IND_NIVEL_HIST"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["FH_FACT"] != System.DBNull.Value)
                {
                    if (Convert.ToInt32(r["FH_FACT"]) > 0)
                    {
                        ff = r["FH_FACT"].ToString();
                        //sb.Append("'").Append(ff.Substring(0, 4)).Append("-")
                        //               .Append(ff.Substring(4, 2)).Append("-").Append(ff.Substring(6, 2));
                        sb.Append(", '").Append(ff).Append("'");
                    }
                    else
                        sb.Append(",null");
                }
                else
                    sb.Append(",null");

                if (r["CD_USUARIO_SCE"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["CD_USUARIO_SCE"].ToString()).Append("'");
                else
                    sb.Append(",null");

                if (r["CD_PROGRAM_SCE"] != System.DBNull.Value)
                    sb.Append(", '").Append(r["CD_PROGRAM_SCE"].ToString()).Append("'");
                else
                    sb.Append(",null");

                sb.Append(", '").Append(Convert.ToDateTime(r["FH_ULT_MODIF_SCE"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("'");
                ff = r["FH_ACT_REGIST_METRA"].ToString();
                sb.Append(", '").Append(ff).Append("'),");
                




                if (totalReg == 250)
                {

                    Console.CursorLeft = 0;
                    Console.Write(registrosLeidos + "/" + total);
                    firstOnly = true;
                    mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                    mcommand = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), mdb.con);
                    mcommand.ExecuteNonQuery();
                    mdb.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    totalReg = 0;
                }

            }
            if (totalReg > 0)
            {
                firstOnly = true;
                Console.CursorLeft = 0;
                Console.Write(registrosLeidos + "/" + total);
                mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.MED);
                mcommand = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), mdb.con);
                mcommand.ExecuteNonQuery();
                mdb.CloseConnection();
                sb = null;
                sb = new StringBuilder();
                totalReg = 0;
            }

            db.CloseConnection();
        }

        private int UltimoDiaCurvasResumen()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            int f = 0;
            f = Convert.ToInt32(DateTime.Now.AddMonths(-1).ToString("yyyyMMdd"));
            
            return f;

        }

        private Dictionary<string, List<EndesaEntity.medida.CurvaResumen>> ConsultaCurvaResumen(List<string> lista_cups20, DateTime fd, DateTime fh, string estado)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            
            StringBuilder sb = new StringBuilder();
            int j = 0;
            DateTime fecha = new DateTime();
            Dictionary<string, List<EndesaEntity.medida.CurvaResumen>> d = new Dictionary<string, List<EndesaEntity.medida.CurvaResumen>>();            

            try
            {
                sb.Append("SELECT r.cd_cups_ext_20, r.cd_cups_ext, r.fh_fact_desde, r.fh_fact_hasta, r.cd_estado_resumen,");
                sb.Append(" r.nm_cons_total_act, r.nm_cons_total_reac, r.nm_curvas_recibidas");


                sb.Append(" FROM med.dt_t_ed_h_rcurvas r WHERE r.cd_cups_ext_20 in (");
                sb.Append("'").Append(lista_cups20[0]).Append("'");
                for (int i = 0; i < lista_cups20.Count; i++)
                    sb.Append(",'").Append(lista_cups20[i]).Append("'");

                sb.Append(") AND r.fh_fact_desde >= ").Append(fd.ToString("yyyyMMdd"));
                sb.Append(" AND r.fh_fact_hasta <= ").Append(fh.ToString("yyyyMMdd")).Append(" and");
                sb.Append(" r.cd_estado_resumen = '").Append(estado).Append("'");

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.CurvaResumen re = new EndesaEntity.medida.CurvaResumen();
                    re.cups22 = r["cd_cups_ext"].ToString();
                    fecha = new DateTime(Convert.ToInt32(r["fh_fact_desde"].ToString().Substring(0, 4)),
                                         Convert.ToInt32(r["fh_fact_desde"].ToString().Substring(4, 2)),
                                         Convert.ToInt32(r["fh_fact_desde"].ToString().Substring(6, 2)));

                    re.fd = fecha;
                    fecha = new DateTime(Convert.ToInt32(r["fh_fact_hasta"].ToString().Substring(0, 4)),
                                         Convert.ToInt32(r["fh_fact_hasta"].ToString().Substring(4, 2)),
                                         Convert.ToInt32(r["fh_fact_hasta"].ToString().Substring(6, 2)));


                    re.fh = fecha;                    
                    re.estado = DescripcionEstadoCurva(r["cd_estado_resumen"].ToString());                    
                    re.activa = Convert.ToInt32(r["nm_cons_total_act"]);
                    re.reactiva = Convert.ToInt32(r["nm_cons_total_reac"]);
                    //if (r["fuente"] != System.DBNull.Value)
                    //    re.fuente = r["fuente"].ToString();

                    List<EndesaEntity.medida.CurvaResumen> rr;
                    if (!d.TryGetValue(re.cups22.Substring(0,20), out rr))
                    {
                        rr = new List<EndesaEntity.medida.CurvaResumen>();
                        rr.Add(re);
                        d.Add(re.cups22.Substring(0, 20), rr);
                    }
                    else
                    {
                        rr.Add(re);
                    }


                }

                db.CloseConnection();


                return d;
                

            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message,
                //"CurvaResumenFunciones.ConsultaCurvaResumen",
                //MessageBoxButtons.OK,
                //MessageBoxIcon.Error);
                Console.WriteLine(e.Message);
                return null;
            }

        }

        private Dictionary<string, List<EndesaEntity.medida.CurvaResumen>> ConsultaCurvaResumen(List<string> lista_cups13, string estado)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            int anio = 0;
            int mes = 0;
            int dia = 0;

            StringBuilder sb = new StringBuilder();
            int j = 0;
            DateTime fecha = new DateTime();
            Dictionary<string, List<EndesaEntity.medida.CurvaResumen>> d = new Dictionary<string, List<EndesaEntity.medida.CurvaResumen>>();

            try
            {
                sb.Append("SELECT r.cd_cups, r.cd_cups_ext_20, r.cd_cups_ext, r.fh_fact_desde, r.fh_fact_hasta, r.cd_estado_resumen,");
                sb.Append(" r.nm_cons_total_act, r.nm_cons_total_reac, r.nm_curvas_recibidas");


                sb.Append(" FROM med.dt_t_ed_h_rcurvas r WHERE r.cd_cups in (");
                sb.Append("'").Append(lista_cups13[0]).Append("'");

                for (int i = 0; i < lista_cups13.Count; i++)
                    sb.Append(",'").Append(lista_cups13[i]).Append("'");

                // sb.Append(") AND r.cd_estado_resumen = '").Append(estado).Append("'");
                sb.Append(") AND r.fh_fact_desde > '" 
                    + DateTime.Now.AddYears(-1).ToString("yyyyMMdd") + "'");

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if(r["cd_estado_resumen"].ToString() == "R")
                    {
                        EndesaEntity.medida.CurvaResumen re = new EndesaEntity.medida.CurvaResumen();
                        re.cups13 = r["cd_cups"].ToString();
                        re.cups22 = r["cd_cups_ext"].ToString();

                        anio = Convert.ToInt32(r["fh_fact_desde"].ToString().Substring(0, 4));
                        mes = Convert.ToInt32(r["fh_fact_desde"].ToString().Substring(4, 2));
                        dia = Convert.ToInt32(r["fh_fact_desde"].ToString().Substring(6, 2));

                        if ((mes >= 1 && mes <= 12) && (dia >= 1 && dia <= 31))
                        {
                            fecha = new DateTime(anio, mes, dia);
                            re.fd = fecha;
                        }
                        else
                            re.fd = DateTime.MinValue;


                        anio = Convert.ToInt32(r["fh_fact_hasta"].ToString().Substring(0, 4));
                        mes = Convert.ToInt32(r["fh_fact_hasta"].ToString().Substring(4, 2));
                        dia = Convert.ToInt32(r["fh_fact_hasta"].ToString().Substring(6, 2));

                        if ((mes >= 1 && mes <= 12) && (dia >= 1 && dia <= 31))
                        {
                            fecha = new DateTime(anio, mes, dia);                            
                            re.fh = fecha;
                        }
                        else
                            re.fh = DateTime.MinValue;





                        re.estado = DescripcionEstadoCurva(r["cd_estado_resumen"].ToString());
                        re.activa = Convert.ToInt32(r["nm_cons_total_act"]);
                        re.reactiva = Convert.ToInt32(r["nm_cons_total_reac"]);
                        //if (r["fuente"] != System.DBNull.Value)
                        //    re.fuente = r["fuente"].ToString();

                        List<EndesaEntity.medida.CurvaResumen> rr;
                        if (!d.TryGetValue(re.cups13, out rr))
                        {
                            rr = new List<EndesaEntity.medida.CurvaResumen>();
                            rr.Add(re);
                            d.Add(re.cups13, rr);
                        }
                        else
                        {
                            rr.Add(re);
                        }
                    }

                    


                }

                db.CloseConnection();


                return d;


            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message,
                //"CurvaResumenFunciones.ConsultaCurvaResumen",
                //MessageBoxButtons.OK,
                //MessageBoxIcon.Error);
                Console.WriteLine(e.Message);
                return null;
            }

        }

        private string DescripcionEstadoCurva(string e)
        {
            string estado = "";

            switch (e)
            {
                case "F":
                    estado = "FACTURADA";
                    break;
                case "R":
                    estado = "REGISTRADA";
                    break;
                default:
                    estado = "SIN CURVA REGISTRADA";
                    break;
            }

            return estado;
        }

        public List<EndesaEntity.medida.CurvaResumen> GetCurvasResumen(string cups20, DateTime fd, DateTime fh)
        {
            List<EndesaEntity.medida.CurvaResumen> l = new List<EndesaEntity.medida.CurvaResumen>();
            List<EndesaEntity.medida.CurvaResumen> o;
            if (dic_curva_resumen.TryGetValue(cups20, out o))
            {
                foreach (EndesaEntity.medida.CurvaResumen p in o)
                {
                    if (p.fd >= fd && p.fh <= fh)
                        l.Add(p);

                }
            }

            return l;
        }

        public EndesaEntity.medida.CurvaResumen Primera_CR_DesdeFecha(string cups13, DateTime fd)
        {
            EndesaEntity.medida.CurvaResumen c;
            List<EndesaEntity.medida.CurvaResumen> o;
            if (dic_curva_resumen.TryGetValue(cups13, out o))
            {
                if (o.Exists(z => z.fd == fd))
                    return o.Find(z => z.fd == fd);
                else
                    return null;
            }
            else
                return null;
            
               
            
        }
    }
}
