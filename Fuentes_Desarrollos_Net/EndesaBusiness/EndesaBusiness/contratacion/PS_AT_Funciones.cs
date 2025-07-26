using EndesaBusiness.global;
using EndesaBusiness.servidores;
using EndesaEntity.cnmc.V21_2019_12_17;
using EndesaEntity.contratacion;
using EndesaEntity;
using EndesaEntity.global;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ff = EndesaBusiness.utilidades;
using MySQLDB = EndesaBusiness.servidores.MySQLDB;
using Microsoft.Graph.TermStore;
using Telegram.Bot.Types;
using EndesaBusiness.medida;

namespace EndesaBusiness.contratacion
{
    public class PS_AT_Funciones : EndesaEntity.contratacion.PS_AT
    {
        public Dictionary<string, EndesaEntity.contratacion.PS_AT> l_ps_at { get; set; }
        private Dictionary<string, string> l_cups;
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "Agora");
        logs.Log ficheroLogPS_AT = new logs.Log(Environment.CurrentDirectory, "logs", "PS_AT");
        utilidades.Param p;
        utilidades.Seguimiento_Procesos ss_pp;

        public PS_AT_Funciones()
        {
            l_ps_at = new Dictionary<string, EndesaEntity.contratacion.PS_AT>();
            l_cups = new Dictionary<string, string>();
            p = new utilidades.Param("ps_param", servidores.MySQLDB.Esquemas.CON);
            ss_pp = new ff.Seguimiento_Procesos();
        }

        public void Crea_PS_AT()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            DateTime ini;
            bool hay_error = false;
            

            try
            {                

                if (HayQueActualizar_PS_AT())
                {

                    ss_pp.Update_Fecha_Inicio("Contratación", "PS_AT", "PS_AT");

                    ini = DateTime.Now;

                    #region Borra tabla ps_at_temp
                    strSql = "delete from ps_at_temp;";
                    Console.WriteLine("Borramos datos de la tabla cont.ps_at_temp");
                    ficheroLogPS_AT.Add("Borramos datos de la tabla cont.ps_at_temp");
                    ficheroLogPS_AT.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    #endregion

                    //Deshabilitamos la actualización de los registros provenientes del SCE, tras migración ya no procede
                    //Anexa_datos_de_PS_Op_a_ps_at_temp();
                    //hay_error = Anexa_datos_gestionPropia_MySQL_ps_at_temp();

                    //ActualizaAutoconsumo();

                    strSql = "REPLACE INTO ps_at_temp"
                        + " SELECT"
                        + " if (ps.cd_empr = 'ENDESA ENERGÍA','EE','EEXXI') AS EMPRESA,"
                        + " ps.cd_cups, ps.de_empr_distdora_nombre,"
                        + " ps.cd_crto_ps AS CCONTATR," //Deberia ser null pero la clave no lo permite y reutilizamos el campo CONTREXT
                        + " if (ps.de_num_ps = 'S/N', NULL, ps.de_num_ps) AS CNUMCATR,"
                        + " DATE_FORMAT(ps.fh_alta_crto, '%Y%m%d') AS fAltaCont,"
                        + " DATE_FORMAT(ps.fh_prev_baja_crto, '%Y%m%d') AS fPrevBajaCont,"
                        + " '003' AS estadoCont,"
                        + " DATE_FORMAT(ps.fh_prev_inicio_crto, '%Y%m%d') AS FPREALTA,"
                        + " ps.cd_tarifa,"
                        + " DATE_FORMAT(ps.fh_prev_inicio_crto, '%Y%m%d') AS FPSERCON,"
                        + " ps.nm_tension_actual AS TENSION,"
                        + " ps.cd_crto_ps AS CONTREXT,"
                        + " LPAD(ps.nm_vers_crto, 3, '0') AS VERSION,"
                        + " (ps.nm_pot_ctatada_1 * 1000) AS VPOTCAL1,"
                        + " (ps.nm_pot_ctatada_2 * 1000) AS VPOTCAL2,"
                        + " 0 AS TDISCHOR,"
                        + " (ps.nm_pot_ctatada_3 * 1000) AS VPOTCAL3,"
                        + " (ps.nm_pot_ctatada_4 * 1000) AS VPOTCAL4,"
                        + " (ps.nm_pot_ctatada_5 * 1000) AS VPOTCAL5,"
                        + " (ps.nm_pot_ctatada_6 * 1000) AS VPOTCAL6,"
                        + " NULL AS TINDGCPY,"
                        + " NULL AS TTICONPS,"
                        + " DATE_FORMAT(ps.fh_baja_crto, '%Y%m%d') AS fBAJACont,"
                        + " 'N' AS segmentoMercado,"
                        + " ps.de_municip AS municipio,"
                        + " ps.de_prov AS provincia,"
                        + " ps.cd_nif_cif_cli AS NIF,"
                        + " ps.tx_apell_cli AS Cliente,"
                        + " NULL AS tipoCli,"
                        + " ps.cups20 AS CUPS20,"
                        + " ps.cd_cups_ext AS CUPS22,"
                        + " if (ps.lg_gestion_propia = 'N',1,2) AS tipoGestionATR,"
                        + " replace(ps.cd_tp_pto_medida, 'Punto de medida tipo ', '') AS TPUNTMED,"
                        + " ps.de_perfil_consumo AS descripcion_autoconsumo,"
                        + " ps.lg_migrado_sap,"
                        + " ps.created_date AS f_ult_mod"
                        + " FROM cont.t_ed_h_ps ps";
                      // 14/10/2024 - GUS: DESHABILITAMOS EL FILTRO QUE EXPORTA SOLO LOS PUNTOS MIGRADOS
                      //  + " WHERE ps.lg_migrado_sap = 'S'";
                   
                    Console.WriteLine("Anexando datos a ps_at_temp de Rosetta");
                    ficheroLogPS_AT.Add("Anexando datos a ps_at_temp de Rosetta");
                    ficheroLogPS_AT.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();


                    #region Borra tabla PS_AT_Ant                    
                    strSql = "delete from PS_AT_Ant;";
                    Console.WriteLine("borrando tabla cont.PS_AT_Ant");
                    ficheroLogPS_AT.Add("borrando tabla cont.PS_AT_Ant");
                    ficheroLogPS_AT.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    #endregion

                    #region Volcamos el contenido de PS_AT a PS_AT_Ant
                    strSql = "replace into PS_AT_Ant select * from PS_AT;";
                    Console.WriteLine("Anexando datos tabla cont.PS_AT a cont.PS_AT_Ant");
                    ficheroLogPS_AT.Add("Anexando datos tabla cont.PS_AT a cont.PS_AT_Ant");
                    ficheroLogPS_AT.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    #endregion

                    #region Borra tabla PS_AT
                    strSql = "delete from PS_AT;";
                    Console.WriteLine("Borramos datos de la tabla cont.PS_AT");
                    ficheroLogPS_AT.Add("Borramos datos de la tabla cont.PS_AT");
                    ficheroLogPS_AT.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    //Console.WriteLine("Borramos datos de la tabla fact.PS_AT");
                    //ficheroLogPS_AT.Add("Borramos datos de la tabla fact.PS_AT");
                    //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    //command = new MySqlCommand(strSql, db.con);
                    //command.ExecuteNonQuery();
                    //db.CloseConnection();
                    #endregion

                    #region Volcamos el contenido de la tabla ps_at_temp a PS_AT

                    strSql = "replace into PS_AT select * from ps_at_temp;";
                    Console.WriteLine("Anexando datos tabla cont.ps_at_temp a cont.PS_AT");
                    ficheroLogPS_AT.Add("Anexando datos tabla cont.ps_at_temp a cont.PS_AT");
                    ficheroLogPS_AT.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    #endregion

                    #region Rellenamos el campo CONTRATO EXTERNO 
                    //Deshabilitamos la actualización de los registros provenientes del SCE, tras migración ya no procede
                    //strSql = " UPDATE PS_AT ps"
                    //     + " INNER JOIN irf_gestionatr atr ON"
                    //     + " atr.cups = ps.IDU"
                    //     + " SET ps.CONTREXT = atr.contratoPS"
                    //     + " WHERE ps.tipoGestionATR = 2 and ps.CONTREXT is null";

                    //Console.WriteLine("Actualizando el campo CONTREXT");
                    //ficheroLogPS_AT.Add("Actualizando el campo CONTREXT");
                    //ficheroLogPS_AT.Add(strSql);
                    //db = new MySQLDB(MySQLDB.Esquemas.CON);
                    //command = new MySqlCommand(strSql, db.con);
                    //command.ExecuteNonQuery();
                    //db.CloseConnection();

                    #endregion

                    #region Rellenamos el campo CONTRATO EXTERNO 

                    strSql = " UPDATE PS_AT ps"
                         + " INNER JOIN t_ed_h_ps atr ON"
                         + " atr.cups20 = ps.CUPS20"
                         + " SET ps.CONTREXT = atr.cd_crto_ps"
                         + " WHERE ps.tipoGestionATR = 2 and ps.CONTREXT is null";

                    Console.WriteLine("Actualizando el campo CONTREXT");
                    ficheroLogPS_AT.Add("Actualizando el campo CONTREXT");
                    ficheroLogPS_AT.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();



                    #endregion

                    // Guardamos historico de la tabla ps_at
                    Actualiza_PS_AT_HIST();

                    #region Publicamos el contenido de la tabla PS_AT a la contrata y al departamento de medida
                    if(p.GetValue("Envio_FTP") == "S")
                        Publica_PS_AT();

                    if (p.GetValue("Envio_SharePoint") == "S")
                        Publica_PS_AT_SharePoint();

                    // Informamos de la finalización del mismo vía eMail
                    //mc = new office.MailCompose(servidores.MySQLDB.Esquemas.CON, "ps_mail", "ps_at_fin");
                    //mc.Send();

                    ActualizaTablaMySQLProcesos();

                    EnviaCorreo();

                    ficheroLogPS_AT.Add("Generando Exportación PS_AT_EXCEL");
                    ficheroLogPS_AT.Add("=================================");

                    Exporta_PS_AT_Excel();

                    ficheroLogPS_AT.Add("Fin Exportación PS_AT_EXCEL");
                    ficheroLogPS_AT.Add("===========================");
                    // ActualizaProcesoAccess(ini);

                    if (p.GetValue("EnviarMail_InformeExcel_Revision_Contrato") == "S")
                        Exporta_PS_AT_Excel_Revision_Contrato();

                    #endregion

                    ss_pp.Update_Fecha_Fin("Contratación", "PS_AT", "PS_AT");

                }
                else
                {

                }


            }
            catch (Exception e)
            {
                ficheroLogPS_AT.AddError("Crea_PS_AT " + e.Message);
            }
        }

        public void Anexa_datos_de_PS_Op_a_ps_at_temp()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            strSql = "replace into ps_at_temp"
                + " select"
                + " if (ps.EMPRESA = '00020', 'EE', if (ps.EMPRESA = '00070', 'EEXXI', 'ERROR')) as EMPRESA,"
                + " ps.IDU, ps.DDISTRIB, ps.CCONTATR, ps.CNUMCATR, ps.FALTACON, ps.FPREVBAJ, ps.TESTCONT, ps.FPREALTA, ps.CTARIFA,"
                + " ps.FPSERCON,  ps.VTENSIOM, ps.CONTREXT, ps.CNUMSCPS, ps.VPOTCAL1, ps.VPOTCAL2, ps.TDISCHOR, ps.VPOTCAL3,"
                + " ps.VPOTCAL4, ps.VPOTCAL5, ps.VPOTCAL6,"
                + " ps.TINDGCPY, ps.TTICONPS, ps.FBAJACON, ps.CSEGMERC, ps.DMUNICIP, ps.DPROVINC, ps.CNIFDNIC, ps.DAPERSOC,"
                + " null, substr(ps.CUPSREE,1,20), ps.CUPSREE, 1 as TipoGestionATR, ps.TPUNTMED,"
                + " null as descripcion_autoconsumo, 'N' as MIGRADO_SAP, now() as f_ult_mod"
                + " from ps_operaciones ps"
                + " inner"
                + " join (select IDU, Max(FALTACON) AS FALTACON from ps_operaciones group by IDU) as pss on"
                + " pss.IDU = ps.IDU and"
                + " pss.FALTACON = ps.FALTACON; ";


            ficheroLogPS_AT.Add("Anexamos datos de ps_operaciones a ps_at_temp");
            ficheroLogPS_AT.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        public bool Anexa_datos_gestionPropia_MySQL_ps_at_temp()
        {
            string strSql = "";
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader rs;
            int x = 0;
            int i = 0;


            try
            {
                strSql = "select g.estadoCont,g.CUPS,g.distribuidora,g.cliente,g.fAltaCont,"
                    + " g.fBajaCont,g.fPrevBajaCont,g.TARIFA,g.tipoCli,g.PROVINCIA,g.MUNICIPIO,"
                    + " g.segmentoMercado,g.NIF,g.CUPS22, 2 AS Expr1, 'EE' AS Expr2 FROM gestionpropiaatr g;";
                ficheroLogPS_AT.Add("Anexamos datos de Gestión propia ATR a ps_at_temp");
                ficheroLogPS_AT.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                rs = command.ExecuteReader();
                while (rs.Read())
                {
                    x++;
                    i++;

                    if (firstOnly)
                    {
                        sb.Append("replace INTO ps_at_temp (EMPRESA,IDU,DDISTRIB,CCONTATR,CNUMCATR,fAltaCont,fPrevBajaCont,");
                        sb.Append("estadoCont,FPREALTA,TARIFA,FPSERCON,TENSION,CONTREXT,Version,");
                        sb.Append("VPOTCAL1,VPOTCAL2,TDISCHOR,VPOTCAL3,VPOTCAL4,VPOTCAL5,VPOTCAL6,");
                        sb.Append("TINDGCPY,TTICONPS,fBajaCont,segmentoMercado,municipio,provincia,NIF,");
                        sb.Append("Cliente,tipoCli,CUPS20,CUPS22,tipoGestionATR,TPUNTMED) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(rs["Expr2"].ToString().Trim()).Append("', ");
                    sb.Append(rs["CUPS"] == System.DBNull.Value ? "null, " : "'" + ff.FuncionesTexto.ArreglaAcentos(rs["CUPS"].ToString()) + "', ");
                    sb.Append(rs["distribuidora"] == System.DBNull.Value ? "null, " : "'" + ff.FuncionesTexto.ArreglaAcentos(rs["distribuidora"].ToString()) + "', ");
                    sb.Append(DateTime.Now.ToString("yyyyMMddHHmmss")).Append(","); // CCONTATR
                    sb.Append("null, ");
                    sb.Append(rs["fAltaCont"] == System.DBNull.Value ? "null, " : ff.FuncionesTexto.ArreglaAcentos(rs["fAltaCont"].ToString()) + ", ");
                    sb.Append(rs["fPrevBajaCont"] == System.DBNull.Value ? "null, " : ff.FuncionesTexto.ArreglaAcentos(rs["fPrevBajaCont"].ToString()) + ", ");
                    sb.Append(rs["estadoCont"] == System.DBNull.Value ? "null, " : "'" + ff.FuncionesTexto.ArreglaAcentos(rs["estadoCont"].ToString()) + "', ");
                    sb.Append("null, ");
                    sb.Append(rs["TARIFA"] == System.DBNull.Value ? "null, " : "'" + ff.FuncionesTexto.ArreglaAcentos(rs["TARIFA"].ToString()) + "', ");
                    sb.Append("null, ");
                    sb.Append("null, ");
                    sb.Append("null, ");
                    sb.Append("null, ");
                    sb.Append("null, ");
                    sb.Append("null, ");
                    sb.Append("null, ");
                    sb.Append("null, ");
                    sb.Append("null, ");
                    sb.Append("null, ");
                    sb.Append("null, ");
                    sb.Append("null, ");
                    sb.Append("null, ");
                    sb.Append("null, ");
                    sb.Append(rs["segmentoMercado"] == System.DBNull.Value ? "null, " : "'" + ff.FuncionesTexto.ArreglaAcentos(rs["segmentoMercado"].ToString()) + "', ");
                    sb.Append(rs["municipio"] == System.DBNull.Value ? "null, " : "'" + ff.FuncionesTexto.ArreglaAcentos(rs["municipio"].ToString()) + "', ");
                    sb.Append(rs["provincia"] == System.DBNull.Value ? "null, " : "'" + ff.FuncionesTexto.ArreglaAcentos(rs["provincia"].ToString()) + "', ");
                    sb.Append(rs["NIF"] == System.DBNull.Value ? "null, " : "'" + ff.FuncionesTexto.ArreglaAcentos(rs["NIF"].ToString()) + "', ");
                    sb.Append(rs["Cliente"] == System.DBNull.Value ? "null, " : "'" + ff.FuncionesTexto.ArreglaAcentos(rs["Cliente"].ToString()) + "', ");
                    sb.Append(rs["tipoCli"] == System.DBNull.Value ? "null, " : "'" + ff.FuncionesTexto.ArreglaAcentos(rs["tipoCli"].ToString()) + "', ");
                    sb.Append(rs["CUPS22"] == System.DBNull.Value ? "null, " : "'" + ff.FuncionesTexto.ArreglaAcentos(rs["CUPS22"].ToString().Substring(0,20)) + "', ");
                    sb.Append(rs["CUPS22"] == System.DBNull.Value ? "null, " : "'" + ff.FuncionesTexto.ArreglaAcentos(rs["CUPS22"].ToString()) + "', ");
                    sb.Append(rs["Expr1"] == System.DBNull.Value ? "null, " : ff.FuncionesTexto.ArreglaAcentos(rs["Expr1"].ToString()) + ", ");
                    sb.Append("null),");

                    if (x == 200)
                    {
                        Console.WriteLine("Guardando " + i + " registros...");
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
                    Console.WriteLine("Guardando " + i + " registros...");
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    x = 0;
                }
                db.CloseConnection();


                // Rellenamos el campo version con datos de Rosetta

                return false;

            }
            catch (Exception e)
            {
                ficheroLogPS_AT.AddError("Anexa_datos_gestionPropia_PS_AT " + e.Message);
                return true;
            }
        }
        private bool HayQueActualizar_PS_AT()
        {
            DateTime fecha_GestionPropiaATR = new DateTime(2010, 01, 01);
            DateTime fecha_PS = new DateTime(2010, 01, 01);
            DateTime fecha_PS_AT = new DateTime(2010, 01, 01);
            DateTime hoy = new DateTime();

            Int32 nfecha_GestionPropiaAtr;
            Int32 nfecha_PS;
            Int32 nfecha_PS_AT;
            Int32 nhoy;

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            try
            {
                hoy = DateTime.Now;

                // Comprobamos fecha Proceso GestionPropiaATR
                strSql = "select fecha from ps_fechas_procesos where"
                    + " proceso = 'gestionATR'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    fecha_GestionPropiaATR = Convert.ToDateTime(reader["fecha"]);
                }

                db.CloseConnection();

                // Comprobamos fecha Proceso PS
                strSql = "select fecha from ps_fechas_procesos where"
                    + " proceso = 'PS'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    fecha_PS = Convert.ToDateTime(reader["fecha"]);
                }
                db.CloseConnection();

                // Comprobamos fecha Proceso PS_AT
                strSql = "select fecha from ps_fechas_procesos where"
                    + " proceso = 'PSAT'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    fecha_PS_AT = Convert.ToDateTime(reader["fecha"]);
                }
                db.CloseConnection();

                nfecha_GestionPropiaAtr = Convert.ToInt32(fecha_GestionPropiaATR.ToString("yyyyMMdd"));
                nfecha_PS = Convert.ToInt32(fecha_PS.ToString("yyyyMMdd"));
                nfecha_PS_AT = Convert.ToInt32(fecha_PS_AT.ToString("yyyyMMdd"));
                nhoy = Convert.ToInt32(hoy.ToString("yyyyMMdd"));


                // GUS - 16-09-2024: cambiamos condición ya que la gestion propia ATR y la PS no se ejecutarán más tras la migración a SAP
                // if ((nfecha_GestionPropiaAtr == nhoy && nfecha_PS == nhoy) && nfecha_PS_AT < nhoy)
                if (nfecha_PS_AT < nhoy)
                    return true;
                else
                    return false;

            }
            catch (Exception e)
            {
                ficheroLogPS_AT.AddError("Comprueba_ResultadoActualizacionTablasBase --> " + e.Message);
                return false;
            }

        }
        private FileInfo Exporta_PS_AT()
        {
            // ftp_PS_AT
            string strSql = "";

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string nombreArchivo = "";
            bool firstOnly = true;
            bool firstOnly2 = true;


            try
            {
                // Borramos las generaciones de fichero anteriores
                string[] listaArchivos = Directory.GetFiles(GetParameter("PS_AT_ruta"), "PS_AT_*");
                for (int i = 0; i < listaArchivos.Length; i++)
                {
                    FileInfo file = new FileInfo(listaArchivos[i]);
                    file.Delete();
                }


                ficheroLogPS_AT.Add("Exportando tabla PS_AT");
                strSql = "Select EMPRESA,IDU,DDISTRIB,CCONTATR,CNUMCATR,fAltaCont,fPrevBajaCont,estadoCont,FPREALTA,TARIFA,"
                    + "FPSERCON,TENSION,CONTREXT,Version,VPOTCAL1,VPOTCAL2,TDISCHOR,"
                    + "VPOTCAL3 , VPOTCAL4, VPOTCAL5, VPOTCAL6, TINDGCPY, TTICONPS, fBajaCont, "
                    + "segmentoMercado, municipio, provincia, NIF, Cliente, tipoCli, CUPS22, tipoGestionATR, TPUNTMED"
                    + " from PS_AT;";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                nombreArchivo = "PS_AT_" + utilidades.Fichero.UltimoDiaHabil_YYMMDD() + ".txt";
                nombreArchivo = GetParameter("PS_AT_ruta") + nombreArchivo;
                StreamWriter sw = new StreamWriter(nombreArchivo, true);
                while (reader.Read())
                {
                    if (firstOnly)
                    {
                        strSql = "EMPRESA|IDU|DDISTRIB|CCONTATR|CNUMCATR|fAltaCont|fPrevBajaCont|estadoCont|FPREALTA|TARIFA|"
                            + "FPSERCON|TENSION|CONTREXT|Version|VPOTCAL1|VPOTCAL2|TDISCHOR|"
                            + "VPOTCAL3|VPOTCAL4| VPOTCAL5| VPOTCAL6| TINDGCPY|TTICONPS|fBajaCont|"
                            + "segmentoMercado|municipio|provincia|NIF|Cliente|tipoCli|CUPS22|tipoGestionATR|TPUNTMED";
                        sw.WriteLine(strSql);
                        firstOnly = false;
                    }
                    firstOnly2 = true;
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        if (firstOnly2)
                        {
                            strSql = reader[i].ToString();
                            firstOnly2 = false;
                        }
                        else
                            strSql = strSql + "|" + reader[i].ToString();
                    }
                    sw.WriteLine(strSql);

                }
                sw.Close();
                db.CloseConnection();

                ficheroLogPS_AT.Add("Fin exportacion.");

                // Finalmente comprimos el archivo
                FileInfo ficheroTxt = new FileInfo(nombreArchivo);
                ff.ZipUnZip zip = new ff.ZipUnZip();
                zip.ComprimirArchivo(nombreArchivo, nombreArchivo.Replace(".txt", ".zip"));
                ficheroTxt.Delete();

                FileInfo archivoZip = new FileInfo(nombreArchivo.Replace(".txt", ".zip"));
                return archivoZip;

            }
            catch (Exception e)
            {
                ficheroLogPS_AT.AddError("Publica_PS_AT: " + e.Message);
                return null;
            }
        }



        public void Publica_PS_AT()
        {
            FileInfo nombre_archivo = Exporta_PS_AT();

            EndesaBusiness.utilidades.UltimateFTP ftp;

            ficheroLogPS_AT.Add("Publicando fichero PS_AT a ftp.");
            //EndesaBusiness.utilidades.Global g = new EndesaBusiness.utilidades.Global();
            //utilidades.Fichero.EjecutaComando(GetParameter("ftp_PS_AT"), utilidades.Fichero.UltimoDiaHabil_YYMMDD());

            ftp = new EndesaBusiness.utilidades.UltimateFTP(
                        p.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                        p.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                        EndesaBusiness.utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_pass", DateTime.Now, DateTime.Now), true),
                        p.GetValue("ftp_port", DateTime.Now, DateTime.Now));


            ftp.Upload(p.GetValue("ruta_destino_FTP", DateTime.Now, DateTime.Now) + nombre_archivo.Name, nombre_archivo.FullName);

            ficheroLogPS_AT.Add("Fin publicacion fichero PS_AT.");

            nombre_archivo.Delete();

            // Pegamos el fichero en la ruta de medida

        }

        public void Publica_PS_AT_SharePoint()
        {
            //office365.MS_Access msAccess = new office365.MS_Access();

            FileInfo archivo_origen = Exporta_PS_AT();
            FileInfo archivo_destino = new FileInfo(p.GetValue("ruta_sharepoint") + archivo_origen.Name);          

            ficheroLogPS_AT.Add("Publicando fichero PS_AT a SharePoint.");


            if (archivo_destino.Exists)
                archivo_destino.Delete();

            archivo_origen.CopyTo(archivo_destino.FullName);


            // Subimos al SharePoint el archivo zip
            //EndesaEntity.cola.ProcesoCola macro = new EndesaEntity.cola.ProcesoCola();
            //macro.ruta = p.GetValue("ruta_access_macro");
            //macro.bbdd = p.GetValue("bbdd_macro_actualiza");
            //macro.nombre_proceso = p.GetValue("macro");
            //macro = msAccess.EjecutaMacro(macro);


            ficheroLogPS_AT.Add("Fin publicacion fichero SharePoint.");

            archivo_origen.Delete();

            // Pegamos el fichero en la ruta de medida

        }

        public void Exporta_PS_AT_Excel_Revision_Contrato()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            bool firstOnly = true;
            int f = 1;
            int c = 1;
            string cups13 = "";

            DateTime fecha = new DateTime();

            try
            {
                EndesaBusiness.utilidades.Fichero.BorrarArchivos_MenosUltimo(p.GetValue("Ubicacion_Informes"), 
                    p.GetValue("PS_AT_Excel_Revision_Contrato"), "xlsx", null);


                string ruta_salida_archivo = p.GetValue("Ubicacion_Informes")
                  + p.GetValue("PS_AT_Excel_Revision_Contrato")
                  + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";


                FileInfo fileInfo = new FileInfo(ruta_salida_archivo);

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("PS_AT");
                var headerCells = workSheet.Cells[1, 1, 1, 6];
                var headerFont = headerCells.Style.Font;

                workSheet.View.FreezePanes(2, 1);

                headerFont.Bold = true;

                strSql = "SELECT ps.Cliente, ps.NIF, ps.IDU, ps.CUPS22, ps.CONTREXT, ps.Version"
                    + " FROM PS_AT ps"
                    + " WHERE ps.estadoCont = '008'";


                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    c = 1;
                    #region Cabecera
                    if (firstOnly)
                    {
                        workSheet.Cells[f, c].Value = "Cliente";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "NIF";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;
                        
                        workSheet.Cells[f, c].Value = "CUPS13";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "CUPS22";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "CONTREXT";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "VERSIÓN";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;



                        firstOnly = false;

                    }
                    #endregion


                    c = 1;
                    f++;

                    if (r["Cliente"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["Cliente"].ToString();
                   
                    c++;

                    if (r["NIF"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["NIF"].ToString();
                    c++;                                       

                    if (r["IDU"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["IDU"].ToString();
                        
                    }

                    c++;

                    if (r["CUPS22"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CUPS22"].ToString();
                    c++;

                    if (r["CONTREXT"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CONTREXT"].ToString();
                    c++;

                    if (r["Version"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["Version"].ToString();
                    c++;

                    

                    

                }
                db.CloseConnection();

                var allCells = workSheet.Cells[1, 1, f, c];
                headerCells = workSheet.Cells[1, 1, 1, c];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;
                allCells = workSheet.Cells[1, 1, f, c];


                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:F1"].AutoFilter = true;
                allCells.AutoFitColumns();



                excelPackage.Save();

                if (p.GetValue("EnviarMail_InformeExcel_Revision_Contrato") == "S")
                    EnvioCorreoInformeExcel_PS_AT_Revision_Contrato(ruta_salida_archivo);


            }
            catch (Exception e)
            {
                ficheroLogPS_AT.AddError("Exporta_PS_AT_Excel en revisión del contrato: " + e.Message);
            }
        }

        public void Exporta_PS_AT_Excel()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            bool firstOnly = true;
            int f = 1;
            int c = 1;
            string cups13 = "";

            DateTime fecha = new DateTime();

            try
            {
                EndesaBusiness.utilidades.Fichero.BorrarArchivos_MenosUltimo(p.GetValue("Ubicacion_Informes"), p.GetValue("PS_AT_Excel"), "xlsx", null);


                string ruta_salida_archivo = p.GetValue("Ubicacion_Informes")
                  + p.GetValue("PS_AT_Excel")
                  + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";


                FileInfo fileInfo = new FileInfo(ruta_salida_archivo);

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("PS_AT");
                var headerCells = workSheet.Cells[1, 1, 1, 25];
                var headerFont = headerCells.Style.Font;

                workSheet.View.FreezePanes(2, 1);

                headerFont.Bold = true;

                strSql = "SELECT ps.EMPRESA, ps.NIF, ps.Cliente,"
                    + " ps.IDU as CUPS13, ps.CUPS22, ps.DDISTRIB AS DISTRIBUIDORA,"
                    + " ps.CCONTATR AS CONTRATO_ATR, CNUMCATR AS NUM_CONTRATO_ATR,"
                    + " ps.fAltaCont AS F_ALTA_CONTRATO, fPrevBajaCont AS F_PREVISTA_BAJA,"
                    + " ps.fBajaCont AS F_BAJA_CONTRATO,"
                    + " ec.Descripcion AS ESTADO_CONTRATO,"
                    + " ps.FPREALTA, ps.TARIFA, ps.FPSERCON AS F_PUESTA_SERVICIO, ps.TENSION,"
                    + " ps.CONTREXT AS CONTRATO_EXTERNO,"
                    + " ps.Version AS VERSION_CONTRATO_EXTERNO,"
                    + " ps.VPOTCAL1 AS POTENCIA_CONTRATADA_P1,"
                    + " ps.VPOTCAL2 AS POTENCIA_CONTRATADA_P2,"
                    + " ps.VPOTCAL3 AS POTENCIA_CONTRATADA_P3,"
                    + " ps.VPOTCAL4 AS POTENCIA_CONTRATADA_P4,"
                    + " ps.VPOTCAL5 AS POTENCIA_CONTRATADA_P5,"
                    + " ps.VPOTCAL6 AS POTENCIA_CONTRATADA_P6,"
                    + " ps.TINDGCPY, tipo_contador.descripcion as TTICONPS, segmentoMercado,"
                    + " ps.municipio, ps.provincia,"
                    + " if (ps.tipoGestionATR = 1, 'No', 'Si') AS GESTION_PROPIA_ATR,"
                    + " ps.TPUNTMED AS TIPO_PUNTO_MEDIDA,"
                    + " ps.descripcion_autoconsumo, ps.f_ult_mod AS F_ULTIMA_ACTUALIZACION"
                    + " FROM cont.PS_AT AS ps"
                    + " LEFT OUTER JOIN cont.cont_estadoscontrato ec ON"
                    + " ec.Cod_Estado = ps.estadoCont"
                    + " LEFT OUTER JOIN cont.cont_ticonsps AS tipo_contador ON"
                    + " tipo_contador.tticonps = ps.TTICONPS"                    
                    + " ORDER BY NIF";


                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    c = 1;
                    #region Cabecera
                    if (firstOnly)
                    {
                        workSheet.Cells[f, c].Value = "EMPRESA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "NIF";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "RAZÓN SOCIAL"; 
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "CUPS13"; 
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "CUPS22"; 
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "TARIFA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "DISTRIBUIDORA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "CONTRATO ATR";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "VERSIÓN CONTRATO ATR";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "FECHA ALTA CONTRATO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "FECHA PRE ALTA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "FECHA PUESTA EN SERVICIO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "FECHA PREVISTA BAJA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "FECHA BAJA CONTRATO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "ESTADO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "TENSIÓN";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "CONTRATO EXTERNO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "VERSIÓN CONT EXTERNO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;



                        for (int i = 1; i <= 6; i++)
                        {
                            workSheet.Cells[f, c].Value = "POTENCIA CONTRATADA P" + i;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); 
                            c++;
                        }


                        workSheet.Cells[f, c].Value = "TINDGCPY";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "TIPO CONTRATO PS";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "SEGMENTO MERCADO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "MUNICIPIO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "PROVINCIA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "GESTIÓN PROPIA ATR";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "TIPO PUNTO MEDIDA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "TIPO AUTOCONSUMO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "ÚLTIMA ACTUALIZACIÓN";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        firstOnly = false;

                    }
                    #endregion


                    c = 1;
                    f++;

                    if (r["EMPRESA"] != System.DBNull.Value)                    
                        workSheet.Cells[f, c].Value = r["EMPRESA"].ToString();                    
                    c++;

                    if (r["NIF"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["NIF"].ToString();
                    c++;

                    if (r["Cliente"] != System.DBNull.Value)                    
                        workSheet.Cells[f, c].Value = r["Cliente"].ToString();
                    
                        
                    c++;

                    if (r["CUPS13"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["CUPS13"].ToString();
                        cups13 = r["CUPS13"].ToString();
                    }
                        
                    c++;

                    if (r["CUPS22"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CUPS22"].ToString();
                    c++;

                    if (r["TARIFA"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                    c++;

                    if (r["DISTRIBUIDORA"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["DISTRIBUIDORA"].ToString();
                    c++;

                    if (r["CONTRATO_ATR"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CONTRATO_ATR"].ToString();
                    c++;

                    if (r["NUM_CONTRATO_ATR"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["NUM_CONTRATO_ATR"]);
                    c++;

                    if (r["F_ALTA_CONTRATO"] != System.DBNull.Value)
                        if((Convert.ToInt32(r["F_ALTA_CONTRATO"]) > 40000000))
                        {
                            fecha = new DateTime(4999, 12, 31);
                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else if (Convert.ToInt32(r["F_ALTA_CONTRATO"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["F_ALTA_CONTRATO"].ToString().Substring(0, 4)),
                                Convert.ToInt32(r["F_ALTA_CONTRATO"].ToString().Substring(4, 2)),
                                Convert.ToInt32(r["F_ALTA_CONTRATO"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    c++;


                    if (r["FPREALTA"] != System.DBNull.Value)
                        if ((Convert.ToInt32(r["FPREALTA"]) > 40000000))
                        {
                            fecha = new DateTime(4999, 12, 31);
                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else if (Convert.ToInt32(r["FPREALTA"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["FPREALTA"].ToString().Substring(0, 4)),
                                Convert.ToInt32(r["FPREALTA"].ToString().Substring(4, 2)),
                                Convert.ToInt32(r["FPREALTA"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    c++;


                    if (r["F_PUESTA_SERVICIO"] != System.DBNull.Value)
                        if ((Convert.ToInt32(r["F_PUESTA_SERVICIO"]) > 40000000))
                        {
                            fecha = new DateTime(4999, 12, 31);
                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else if (Convert.ToInt32(r["F_PUESTA_SERVICIO"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["F_PUESTA_SERVICIO"].ToString().Substring(0, 4)),
                                Convert.ToInt32(r["F_PUESTA_SERVICIO"].ToString().Substring(4, 2)),
                                Convert.ToInt32(r["F_PUESTA_SERVICIO"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    c++;



                    if (r["F_PREVISTA_BAJA"] != System.DBNull.Value)
                        if ((Convert.ToInt32(r["F_PREVISTA_BAJA"]) > 40000000))
                        {
                            fecha = new DateTime(4999, 12, 31);
                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else if (Convert.ToInt32(r["F_PREVISTA_BAJA"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["F_PREVISTA_BAJA"].ToString().Substring(0, 4)),
                                Convert.ToInt32(r["F_PREVISTA_BAJA"].ToString().Substring(4, 2)),
                                Convert.ToInt32(r["F_PREVISTA_BAJA"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    c++;

                    if (r["F_BAJA_CONTRATO"] != System.DBNull.Value)
                        if ((Convert.ToInt32(r["F_BAJA_CONTRATO"]) > 40000000))
                        {
                            fecha = new DateTime(4999, 12, 31);
                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        else if (Convert.ToInt32(r["F_BAJA_CONTRATO"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["F_BAJA_CONTRATO"].ToString().Substring(0, 4)),
                                Convert.ToInt32(r["F_BAJA_CONTRATO"].ToString().Substring(4, 2)),
                                Convert.ToInt32(r["F_BAJA_CONTRATO"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    c++;
                    

                    if (r["ESTADO_CONTRATO"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["ESTADO_CONTRATO"].ToString();
                    c++;

                    if (r["TENSION"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["TENSION"].ToString();
                    c++;

                    if (r["CONTRATO_EXTERNO"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CONTRATO_EXTERNO"].ToString();
                    c++;

                    if (r["VERSION_CONTRATO_EXTERNO"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["VERSION_CONTRATO_EXTERNO"]);
                    c++;

                    for(int i = 1; i <= 6; i++)
                    {
                        if (r["POTENCIA_CONTRATADA_P" + i] != System.DBNull.Value)
                        {
                            if (Convert.ToInt32(r["POTENCIA_CONTRATADA_P" + i]) > 0)
                            {
                                workSheet.Cells[f, c].Value = Convert.ToDouble(r["POTENCIA_CONTRATADA_P" + i]);
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                c++;
                            }
                            else
                                c++;
                                
                        }
                        else
                            c++;                        
                    }

                    if (r["TINDGCPY"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["TINDGCPY"].ToString();
                    c++;

                    if (r["TTICONPS"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["TTICONPS"].ToString();
                    c++;

                    if (r["segmentoMercado"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["segmentoMercado"].ToString();
                    c++;

                    if (r["municipio"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["municipio"].ToString();
                    c++;

                    if (r["provincia"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["provincia"].ToString();
                    c++;

                    if (r["GESTION_PROPIA_ATR"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["GESTION_PROPIA_ATR"].ToString();
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                        
                    c++;

                    if (r["TIPO_PUNTO_MEDIDA"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["TIPO_PUNTO_MEDIDA"]);
                    c++;
                    
                    if (r["descripcion_autoconsumo"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["descripcion_autoconsumo"].ToString();
                    c++;

                    if (r["F_ULTIMA_ACTUALIZACION"] != System.DBNull.Value)
                    {                       

                        workSheet.Cells[f, c].Value = Convert.ToDateTime(r["F_ULTIMA_ACTUALIZACION"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                        
                    c++;

                }
                db.CloseConnection();

                var allCells = workSheet.Cells[1, 1, f, c];
                headerCells = workSheet.Cells[1, 1, 1, c];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;
                allCells = workSheet.Cells[1, 1, f, c];
                               

                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:AG1"].AutoFilter = true;
                allCells.AutoFitColumns();

                

                excelPackage.Save();

                if (p.GetValue("EnviarMail_InformeExcel") == "S")
                    EnvioCorreoInformeExcel_PS_AT(ruta_salida_archivo);


            }
            catch(Exception e)
            {
                ficheroLogPS_AT.AddError("Exporta_PS_AT_Excel: " + e.Message);
            }
        }

        

        private void EnvioCorreoInformeExcel_PS_AT(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                Console.WriteLine("Enviando correo Informe PS_AT vía Excel");

                string from = p.GetValue("Buzon_envio_email");
                string to = p.GetValue("email_PS_AT_Excel_para");
                string cc = p.GetValue("email_PS_AT_Excel_copia");
                string subject = "Exportación datos PS_AT a " + " " + DateTime.Now.ToString("dd/MM/yyyy");
                //string body = GeneraCuerpoHTML(CreaTabla(false), "No &Aacute;gora");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   Adjuntamos el informe Excel de la extracción PS_AT.");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");


                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (p.GetValue("EnviarMail_InformeExcel") == "S")
                    mes.SendMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);

                else
                    mes.SaveMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);

                ficheroLog.Add("Correo enviado desde: " + p.GetValue("Buzon_envio_email"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Envio Mail Excel PS_AT: " + e.Message);
            }
        }

        private void EnvioCorreoInformeExcel_PS_AT_Revision_Contrato(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                Console.WriteLine("Enviando correo Informe PS_AT vía Excel revisión del contrato");

                string from = p.GetValue("Buzon_envio_email");
                string to = p.GetValue("email_PS_AT_Excel_Revision_Contrato_para");
                string cc = p.GetValue("email_PS_AT_Excel_Revision_Contrato_copia");
                string subject = "Exportación datos PS_AT en estado revisión del contrato a " + " " + DateTime.Now.ToString("dd/MM/yyyy");
                //string body = GeneraCuerpoHTML(CreaTabla(false), "No &Aacute;gora");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("   Adjuntamos el informe Excel de la extracción PS_AT,");
                textBody.Append(" en estado 'Revisión del contrato'.");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");


                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (p.GetValue("EnviarMail_InformeExcel_Revision_Contrato") == "S")
                    mes.SendMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);

                else
                    mes.SaveMail(from, to, cc, subject, utilidades.FuncionesTexto.TextToHtml(textBody.ToString()), archivo);

                ficheroLog.Add("Correo enviado desde: " + p.GetValue("Buzon_envio_email"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Envio Mail Excel PS_AT en revisión del contrato: " + e.Message);
            }
        }

        private void Actualiza_PS_AT_HIST()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            try
            {              

                //strSql = "replace into PS_AT_HIST select PS_AT.*, substr(PS_AT.CUPS22,1,20), date_format(now(),'%Y-%m-%d') Fecha_Anexion from PS_AT;";
                strSql = "REPLACE INTO PS_AT_HIST SELECT EMPRESA, IDU, DDISTRIB, CCONTATR, CNUMCATR, fAltaCont,"
                    + " fPrevBajaCont, estadoCont, FPREALTA, TARIFA, FPSERCON, TENSION, CONTREXT, Version,"
                    + " VPOTCAL1, VPOTCAL2, TDISCHOR, VPOTCAL3, VPOTCAL4, VPOTCAL5, VPOTCAL6, TINDGCPY, TTICONPS,"
                    + " fBajaCont, segmentoMercado, municipio, provincia, NIF, Cliente, tipoCli, CUPS22, tipoGestionATR,"
                    + " TPUNTMED, descripcion_autoconsumo, MIGRADO_SAP, f_ult_mod, CUPS20, date_format(now(),'%Y-%m-%d') FROM cont.PS_AT;";

                Console.WriteLine(strSql);
                ficheroLogPS_AT.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                ficheroLogPS_AT.AddError("Actualiza_PS_AT_Hist " + e.Message);
            }
        }
        private void ActualizaAutoconsumo()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            try
            {

                strSql = "UPDATE ps_at_temp ps"
                    + " INNER JOIN ps_autoconsumos t ON"
                    + " t.cups22 = ps.CUPS22 AND"
                    + " t.version = ps.Version"
                    + " SET ps.descripcion_autoconsumo = t.desc_complemento"
                    + " WHERE ps.MIGRADO_SAP = 'N'";
                Console.WriteLine(strSql);
                ficheroLogPS_AT.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                strSql = "UPDATE ps_at_temp ps"
                    + " INNER JOIN ps_autoconsumos t ON"
                    +" t.cups22 = ps.CUPS22"
                    + " SET ps.descripcion_autoconsumo = t.desc_complemento"
                    + " WHERE ps.MIGRADO_SAP = 'N' and ps.tipoGestionATR = 2";
                Console.WriteLine(strSql);
                ficheroLogPS_AT.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                               

            }
            catch (Exception e)
            {
                ficheroLogPS_AT.AddError("ActualizaAutoconsumo " + e.Message);
            }

        }

        private void ActualizaTablaMySQLProcesos()
        {

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            DateTime ahora = new DateTime();

            // Actualizamos campos descripcion_autoconsumo de la tabla PS
            ahora = DateTime.Now;
            strSql = "update ps_fechas_procesos set fecha = '" + ahora.ToString("yyyy-MM-dd") + "'"
                + " where proceso = 'PSAT'";

            Console.WriteLine("Actualizamos la fecha de los procesos de la tabla en MySQL");
            ficheroLogPS_AT.Add("ctualizamos la fecha de los procesos de la tabla en MySQL");
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        public void Carga_PS_AT()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            try
            {
                CargaListaCups();

                ficheroLog.Add("Cargando diccionario PS_AT");
                strSql = "select ps.EMPRESA, ps.IDU, ps.CUPS22, ps.NIF, ps.Cliente, ps.TARIFA, ps.provincia, ec.Descripcion ps.estado_contrato"
                    + " from cont.PS_AT ps"
                    + " left outer join cont_estadoscontrato ec on"
                    + " ec.Cod_Estado = ps.estadoCont"
                    + " group by IDU;";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    EndesaEntity.contratacion.PS_AT c = new EndesaEntity.contratacion.PS_AT();
                    c.cups13 = reader["IDU"].ToString();
                    if (reader["CUPS22"] != System.DBNull.Value)
                    {
                        c.cups22 = reader["CUPS22"].ToString();
                        c.cups20 = c.cups22.Substring(0, 20);
                    }

                    if (reader["NIF"] != System.DBNull.Value)
                        c.nif = reader["NIF"].ToString();
                    c.nombre_cliente = reader["Cliente"].ToString();
                    if (reader["TARIFA"] != System.DBNull.Value)
                        c.tarifa = reader["TARIFA"].ToString();
                    if (reader["provincia"] != System.DBNull.Value)
                        c.provincia = reader["provincia"].ToString();
                    if (reader["estado_contrato"] != System.DBNull.Value)
                        c.estado_contrato = reader["estado_contrato"].ToString();

                    if (reader["EMPRESA"] != System.DBNull.Value)
                        c.empresa = reader["EMPRESA"].ToString();

                    if (c.cups22 == null || c.cups22 == "")
                        c.cups22 = GetInfoCups(c.cups13);
                    if (c.cups22 != "")
                        l_ps_at.Add(c.cups22.Substring(0, 20), c);
                }

                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("PS_AT_Funciones.Carga_PS_AT: " + e.Message);
            }
        }
        public void Carga_PS_AT(string empresa)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            try
            {
                CargaListaCups();

                ficheroLog.Add("Cargando diccionario PS_AT");
                strSql = "select IDU, CUPS22, NIF, Cliente, TARIFA, provincia, ec.Descripcion estado_contrato"
                    + " from cont.PS_AT ps"
                    + " left outer join cont_estadoscontrato ec on"
                    + " ec.Cod_Estado = ps.estadoCont"
                    + " where ps.EMPRESA = '" + empresa + "'"
                    + " group by IDU;";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    EndesaEntity.contratacion.PS_AT c = new EndesaEntity.contratacion.PS_AT();
                    c.cups13 = reader["IDU"].ToString();
                    if (reader["CUPS22"] != System.DBNull.Value)
                        c.cups22 = reader["CUPS22"].ToString();
                    if (reader["NIF"] != System.DBNull.Value)
                        c.nif = reader["NIF"].ToString();
                    c.nombre_cliente = reader["Cliente"].ToString();
                    if (reader["TARIFA"] != System.DBNull.Value)
                        c.tarifa = reader["TARIFA"].ToString();
                    if (reader["provincia"] != System.DBNull.Value)
                        c.provincia = reader["provincia"].ToString();
                    if (reader["estado_contrato"] != System.DBNull.Value)
                        c.estado_contrato = reader["estado_contrato"].ToString();

                    if (c.cups22 == null || c.cups22 == "")
                        c.cups22 = GetInfoCups(c.cups13);
                    if (c.cups22 != "")
                        l_ps_at.Add(c.cups22.Substring(0, 20), c);
                }

                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("PS_AT_Funciones.Carga_PS_AT: " + e.Message);
            }
        }

        private string GetParameter(string codigo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            string vcodigo;
            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                strSql = "Select valor from cont.ps_parametros where"
                    + " codigo = '" + codigo + "'";

                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    vcodigo = reader["valor"].ToString();
                }
                else
                {
                    Console.WriteLine("El valor " + codigo + " no está parametrizado en cont.ps_parametros.");
                    vcodigo = "";
                }
                db.CloseConnection();
                return vcodigo;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLogPS_AT.AddError("GetParameter: " + e.Message);
            }
            return "";

        }

        private void CargaListaCups()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            try
            {
                ficheroLog.Add("Cargando listado cups de med.dt_cups");
                strSql = "select substr(cups15,1,13) as cups13, substr(cups22,1,20) as cups20 from med.dt_cups group by substr(cups15,1,13)";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    string a;
                    if (r["cups13"] != System.DBNull.Value && r["cups20"] != System.DBNull.Value)
                        if (!l_cups.TryGetValue(r["cups13"].ToString(), out a))
                            l_cups.Add(r["cups13"].ToString(), r["cups20"].ToString());
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("CargaListaCups: " + e.Message);
            }
        }

        private string GetInfoCups(string cups13)
        {
            string a;
            if (l_cups.TryGetValue(cups13, out a))
                return a;
            else
                return "";

        }

        public bool Existe_CUPS20(string cups20)
        {
            EndesaEntity.contratacion.PS_AT o;
            if (l_ps_at.TryGetValue(cups20, out o))
            {
                this.nif = o.nif;
                this.nombre_cliente = o.nombre_cliente;
                this.empresa = o.empresa;
                this.estado_contrato = o.estado_contrato;
                return true;
            }
            else
                return false;
        }

        public void EnviaCorreo()
        {

            string body = "";
            string subject = "";
            string from = "";
            string to = "";
            string cc = null;
            string adjuntos = null;
            string[] listaArchivos;
            FileInfo file;
            bool firstOnly = true;

            EndesaBusiness.utilidades.ListasCorreo listasCorreo =
                new EndesaBusiness.utilidades.ListasCorreo(servidores.MySQLDB.Esquemas.CON, "ps_mail", "ps_at_fin");


            body = listasCorreo.correo.body;

            from = listasCorreo.correo.mailbox;
            to = listasCorreo.correo.to;
            cc = listasCorreo.correo.cc;
            subject = listasCorreo.correo.subject;

            //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
            EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

            mes.SendMail(from, to, cc, subject, body, adjuntos);
            Thread.Sleep(1000);
            mes = null;
            

        }
    }
}
