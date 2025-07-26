using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class Alarmas
    {
        private Dictionary<string, EndesaEntity.Alarma> dic_alarma;
        private EndesaBusiness.utilidades.Param param;
        logs.Log ficheroLog;

        public Alarmas()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Alarmas");
            param = new utilidades.Param("fact_alarmas_param", MySQLDB.Esquemas.FAC);
            dic_alarma = new Dictionary<string, EndesaEntity.Alarma>();
        }

        public void ImportarAlarmas()
        {

            string md5 = "";
            string archivo_alarmas = "";
            string archivo_comentarios = "";
            
            try
            {                               

                ficheroLog.Add("Ejecutando extractor: "  + param.GetValue("extractor"));
                Console.WriteLine("Ejecutando extractor: " + param.GetValue("extractor"));
                utilidades.Fichero.EjecutaComando(param.GetValue("extractor"));
                ficheroLog.Add("Fin extractor");

                
                archivo_alarmas = param.GetValue("inbox") + param.GetValue("archivo_alarmas");
                md5 = utilidades.Fichero.checkMD5(archivo_alarmas).ToString();
                ficheroLog.Add("Comparando MD5 alarmas: " + md5 + " vs " + param.GetValue("MD5_archivo_alarmas"));
                Console.WriteLine("Comparando MD5 alarmas: " + md5 + " vs " + param.GetValue("MD5_archivo_alarmas"));

                if (md5 != param.GetValue("MD5_archivo_alarmas"))
                {

                    ImportaAlarmas(archivo_alarmas,md5);
                  
                }

                archivo_comentarios = param.GetValue("inbox") + param.GetValue("archivo_comentarios");
                md5 = utilidades.Fichero.checkMD5(archivo_comentarios).ToString();

                ficheroLog.Add("Comparando MD5 comentarios: " + md5 + " vs " + param.GetValue("MD5_archivo_alarmas"));
                Console.WriteLine("Comparando MD5 comentarios: " + md5 + " vs " + param.GetValue("MD5_archivo_alarmas"));

                if (md5 != param.GetValue("MD5_archivo_comentarios"))
                {

                    ImportaComentarios(archivo_comentarios, md5);
                    
                }

            }
            catch (Exception e)
            {
                ficheroLog.AddError(e.Message);
            }
        }

        private void ImportaAlarmas(string archivo, string md5)
        {
            long i = 0;
            long j = 0;
            long total_lineas = 0;
            System.IO.StreamReader file;
            string linea = "";
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            MySQLDB db;
            MySqlCommand command;
            string[] c;

            try
            {
                ficheroLog.Add("Cargando Alarmas: " + archivo);
                Console.WriteLine("Cargando Alarmas: " + archivo);
                Console.WriteLine("");
                file = new System.IO.StreamReader(archivo, System.Text.Encoding.GetEncoding(1252));
                while ((linea = file.ReadLine()) != null)
                {

                    i = 1;
                    c = linea.Split(';');
                    j++;
                    total_lineas++;

                    if (firstOnly)
                    {
                        firstOnly = false;
                        sb.Append("replace into fact_alarmas ");
                        sb.Append("(CCOUNIPS,CONTREXT,CCONTRPS,CNUMSCPS,TESTCONT,TTICONPS,TUSOCPS,");
                        sb.Append("TINDGCPY,TCONTSVA,FINIALAR,FFINALAR,AlarmaID,UserID) values ");

                    }
                    sb.Append("('").Append(c[i]).Append("',"); i++; // CCOUNIPS
                    sb.Append("'").Append(c[i]).Append("',"); i++; // CONTREXT
                    sb.Append("'").Append(c[i]).Append("',"); i++; // CCONTRPS
                    sb.Append(c[i]).Append(","); i++; // CNUMSCPS
                    sb.Append(c[i]).Append(","); i++; // TESTCONT
                    sb.Append("'").Append(c[i]).Append("',"); i++; // TTICONPS
                    sb.Append("'").Append(c[i]).Append("',"); i++; // TUSOCPS
                    sb.Append("'").Append(c[i]).Append("',"); i++; // TINDGCPY
                    sb.Append("'").Append(c[i]).Append("',"); i++; // TCONTSVA

                    if (c[i] != "00000000") // FINIALAR
                        sb.Append("'").Append(c[i].Substring(0, 4)).Append("-")
                            .Append(c[i].Substring(4, 2)).Append("-").Append(c[i].Substring(6, 2)).Append("',");
                    else
                        sb.Append("null,");
                    i++;

                    if (c[i] != "00000000") // FFINALAR
                        sb.Append("'").Append(c[i].Substring(0, 4)).Append("-")
                            .Append(c[i].Substring(4, 2)).Append("-").Append(c[i].Substring(6, 2)).Append("',");
                    else
                        sb.Append("null,");

                    i++;

                    //sb.Append(c[i]).Append(","); i++; // AlarmaID
                    //sb.Append(c[i]).Append("),"); i++; // UserID

                    sb.Append("null, 1),");

                    if (j == 250)
                    {
                        Console.CursorLeft = 0;
                        Console.Write("Leyendo " + total_lineas + " lineas...");
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        j = 0;
                    }
                }


                if (j > 0)
                {
                    ficheroLog.Add("Leyendo " + total_lineas + " lineas...");
                    Console.WriteLine("Leyendo " + total_lineas + " lineas...");
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    j = 0;
                }

                file.Close();
                param.code = "MD5_archivo_alarmas";
                param.from_date = new DateTime(2021, 12, 31);
                param.to_date = new DateTime(4999, 12, 31);
                param.value = md5;
                param.Save();

            }
            catch(Exception e)
            {
                ficheroLog.AddError("CargaAlarmas; " + e.Message);
            }
            
        }
        private void ImportaComentarios(string archivo, string md5)
        {
            long total_lineas = 0;
            long i = 0;
            long j = 0;
            System.IO.StreamReader file;
            string linea = "";
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            MySQLDB db;
            MySqlCommand command;
            string[] c;

            try
            {
                
                ficheroLog.Add("Cargando Comentarios: " + archivo);
                Console.WriteLine("Cargando Comentarios: " + archivo);
                Console.WriteLine("");
                file = new System.IO.StreamReader(archivo, System.Text.Encoding.GetEncoding(1252));
                while ((linea = file.ReadLine()) != null)
                {

                    i = 1;
                    c = linea.Split(';');
                    j++;
                    total_lineas++;

                    if (firstOnly)
                    {
                        firstOnly = false;
                        sb.Append("replace into fact_alarmas_comentarios ");
                        sb.Append("(TOBJETO,EMPRESA,CLINNEG,CCONTRAT,FCOMENT,CUSERID,DCOMENTA,");
                        sb.Append("UserID) values ");

                    }
                    sb.Append("(").Append(c[i]).Append(","); i++; // TOBJETO
                    sb.Append("'").Append(c[i]).Append("',"); i++; // EMPRESA                    
                    sb.Append(c[i]).Append(","); i++; // CLINNEG
                    sb.Append("'").Append(c[i]).Append("',"); i++; // CCONTRAT

                    if (c[i] != "00000000") // FCOMENT
                        sb.Append("'").Append(c[i].Substring(0, 4)).Append("-")
                            .Append(c[i].Substring(4, 2)).Append("-").Append(c[i].Substring(6, 2)).Append("',");
                    else
                        sb.Append("null,");

                    i++;


                    sb.Append("'").Append(c[i]).Append("',"); i++; // CUSERID
                    sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(c[i])).Append("',"); i++; // DCOMENTA
                    sb.Append("1),"); i++; // UserID
                   

                    if (j == 250)
                    {
                        Console.CursorLeft = 0;
                        Console.Write("Leyendo " + total_lineas + " lineas...");
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        j = 0;
                    }
                }


                if (j > 0)
                {
                    ficheroLog.Add("Leyendo " + total_lineas + " lineas...");
                    Console.CursorLeft = 0;
                    Console.WriteLine("Leyendo " + total_lineas + " lineas...");

                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    j = 0;
                }

                file.Close();

                param.code = "MD5_archivo_comentarios";
                param.from_date = new DateTime(2021, 12, 31);
                param.to_date = new DateTime(4999, 12, 31);
                param.value = md5;
                param.Save();

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Carga Comentarios: " + e.Message);
            }

        }

        public void Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime maxFecha = new DateTime();

            try
            {
                strSql = "select max(F_ULT_MOD) maxFecha from fact_alarmas";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    if (r["maxFecha"] != System.DBNull.Value)
                        maxFecha = Convert.ToDateTime(r["maxFecha"]);
                }
                db.CloseConnection();

                strSql = "select a.CCOUNIPS, r.CUPS20, a.CONTREXT, a.CCONTRPS, a.CNUMSCPS, a.TESTCONT, a.FINIALAR, a.FFINALAR,"
                    + " c.dcomenta , a.F_ULT_MOD"
                    + " from fact_alarmas a"
                    + " left outer join fact_alarmas_comentarios c on"
                    + " c.CCONTRAT = a.ccontrps"
                    + " inner join relacion_cups r on"
                    + " r.CUPS_CORTO = a.CCOUNIPS"
                    + " where a.F_ULT_MOD >= '" + maxFecha.ToString("yyyy-MM-dd") + "' and"
                    + " a.TESTCONT <> 6";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    EndesaEntity.Alarma c = new EndesaEntity.Alarma();

                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        c.cups13 = r["CCOUNIPS"].ToString();

                    if (r["CUPS20"] != System.DBNull.Value)
                        c.cups20 = r["CUPS20"].ToString();

                    if (r["CONTREXT"] != System.DBNull.Value)
                        c.contrext = r["CONTREXT"].ToString();

                    if (r["CCONTRPS"] != System.DBNull.Value)
                        c.ccontrps = r["CCONTRPS"].ToString();

                    if (r["CNUMSCPS"] != System.DBNull.Value)
                        c.cnumscps = Convert.ToInt32(r["CNUMSCPS"]);

                    if (r["TESTCONT"] != System.DBNull.Value)
                        c.testcont = Convert.ToInt32(r["TESTCONT"]);

                    if (r["FINIALAR"] != System.DBNull.Value)
                        c.finialar = Convert.ToDateTime(r["FINIALAR"]);

                    if (r["FFINALAR"] != System.DBNull.Value)
                        c.ffinalar = Convert.ToDateTime(r["FFINALAR"]);

                    if (r["dcomenta"] != System.DBNull.Value)
                        c.dcomenta = r["dcomenta"].ToString();

                    EndesaEntity.Alarma a;
                    if (!dic_alarma.TryGetValue(c.cups20, out a))
                        dic_alarma.Add(c.cups20, c);

                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void CargaSAP()
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql;
            
            DateTime maxFecha = new DateTime();

            try
            {
                strSql = "select trunc(max(a.fec_act)) as maxFecha from ed_owner.t_ed_d_alarmasconfact a;";
                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();


                while (r.Read())
                {
                    if (r["maxFecha"] != System.DBNull.Value)
                        maxFecha = Convert.ToDateTime(r["maxFecha"]);
                }
                db.CloseConnection();

                
                strSql = "select a.ccontrps, (case when u.cd_cups_gas_ext is null then u.cd_cups20_metra else u.cd_cups_gas_ext end) as cups20, a.dcomenta, a.finialar, a.ffinalar"
                    + " from ed_owner.t_ed_d_alarmasconfact a  inner join ed_owner.t_ed_f_uvcrtos u on a.ccontrps = u.id_crto_ext"
                    + " where a.fec_act >= '" + maxFecha.ToString("yyyy-MM-dd") + "' and a.ffinalar >=" + maxFecha.ToString("yyyyMMdd") + " and a.dcomenta is not null and cups20 is not null and u.de_estado_crto <> 'BAJA';";

                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    EndesaEntity.Alarma c = new EndesaEntity.Alarma();

                    if (r["cups20"] != System.DBNull.Value)
                        c.cups20 = r["cups20"].ToString();

                    if (r["ccontrps"] != System.DBNull.Value)
                        c.ccontrps = r["ccontrps"].ToString();

                    if (r["finialar"] != System.DBNull.Value)
                        c.finialar = DateTime.ParseExact(r["finialar"].ToString(), "yyyyMMdd", null);  //Convert.ToDateTime(r["finialar"]);

                    if (r["ffinalar"] != System.DBNull.Value)
                        c.ffinalar = DateTime.ParseExact(r["ffinalar"].ToString(), "yyyyMMdd", null); //Convert.ToDateTime(r["ffinalar"]);

                    if (r["dcomenta"] != System.DBNull.Value)
                        c.dcomenta = r["dcomenta"].ToString();

                    EndesaEntity.Alarma a;
                    if (!dic_alarma.TryGetValue(c.cups20, out a))
                        dic_alarma.Add(c.cups20, c);

                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public string GetAlarma(string cusp20)
        {
            string comentario_alarma = "";
            EndesaEntity.Alarma a;
            if (dic_alarma.TryGetValue(cusp20, out a))
                comentario_alarma = a.dcomenta;
            else
                comentario_alarma = "";

            return comentario_alarma;
        }

        public string GetAlarmaSAP(string cups20)
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql;
            string alarma = "";

            try
            {
                strSql = "select a.dcomenta as alarma from ed_owner.t_ed_d_alarmasconfact a"
                    + " inner join ed_owner.t_ed_f_uvcrtos u on a.ccontrps = u.id_crto_ext"
                    + " where u.cd_cups20_metra = '" + cups20 + "' and a.fec_act >= (select current_date-1);";

                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())

                {
                    if (r["alarma"] != System.DBNull.Value)
                        alarma = r["alarma"].ToString();
                    break;
                }
                db.CloseConnection();
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("Error en función GetAlarmaSAP para el CUPS " + cups20 + ": " + ex.Message);
            }
            
            return alarma;
        }
    }
}
