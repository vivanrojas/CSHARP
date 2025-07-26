using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaEntity.medida;

namespace EndesaBusiness.medida
{
    public class ImportacionTPLs
    {
        logs.Log ficheroLog;
        utilidades.Param p;
        public List<PO1011> l_p;
        string prefijo_tabla;        

        public ImportacionTPLs()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_tpls");
            p = new utilidades.Param("tpls_param", servidores.MySQLDB.Esquemas.MED);
            prefijo_tabla = p.GetValue("prefijo_tabla", DateTime.Now, DateTime.Now);
            prefijo_tabla = prefijo_tabla.Trim();
        }

        public ImportacionTPLs(bool medida_cuarto_horaria)
        {
            PS ps = new PS();
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            bool firstOnly = true;
            bool hayError = false;

            l_p = new List<PO1011>();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "tpls");
            p = new utilidades.Param("tpls_param", servidores.MySQLDB.Esquemas.MED);

            try
            {

                ficheroLog.Add("======================================");
                ficheroLog.Add("*********** INICIO PROCESO ***********");
                ficheroLog.Add("======================================");


                prefijo_tabla = p.GetValue("prefijo_tabla", DateTime.Now, DateTime.Now);
                prefijo_tabla = prefijo_tabla.Trim();

                // comprobamos que la carpeta de errores en fichero exista
                DirectoryInfo dir = new DirectoryInfo(p.GetValue("ruta_archivos_error", DateTime.Now, DateTime.Now));
                if (!dir.Exists)
                    dir.Create();

                // *** Se unifican los directorios en uno solo *** // 
                //ProcesaDirectorio("ruta_curvas_ftp", "curvas_ftp");
                //ProcesaDirectorio("ruta_fenosa", "fenosa");
                //ProcesaDirectorio("ruta_viesgo", "viesgo");

                hayError = ProcesaDirectorio(medida_cuarto_horaria ? "inbox_cuartohoraria" : "inbox", "todos");



                SinValores_AE_R1(); // ---> Guarda en una tabla los valores nulos que han llegado en fichero
                ANEXA_TPL_A_CCHODLAST();
                ArreglaUltimoDomingoMarzo(medida_cuarto_horaria);
                ps.Update_ps_mstr(); // --> Recorre la tabla cchodlast para reconstruir los IDMP de la tabla ps_mstr.
                ps.Carga(); // --> Una vez cargada cargamos en memoria todos los IDMP para el proceso posterior.

                // Para dividir el proceso realizamos el volcado de CCHODLAST a HISTORICO CUPS a CUPS

                ficheroLog.Add("Inicio CCHODLAST_A_CCHODLAST_HISTORICO");
                ficheroLog.Add("======================================");
                strSql = "delete from " + prefijo_tabla + "med_cc_horaria_odhistorico_temporal";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "select CUPS22 from " + prefijo_tabla + "cchodlast group by CUPS22";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (firstOnly)
                    {
                        ficheroLog.Add("Volcado_de_cchodlast_A_med_cc_horaria_odhistorico_tmp");
                        ficheroLog.Add("=====================================================");
                        firstOnly = false;
                    }
                    Volcado_de_cchodlast_A_med_cc_horaria_odhistorico_tmp(r["CUPS22"].ToString(), ps);
                }

                db.CloseConnection();

                Volcado_de_med_cc_horaria_odhistorico_tmp_A_med_cc_horaria_odhistorico();
                Borrar_tabla(prefijo_tabla + "cchodlast");
                Borrar_tabla("tpls");

                ficheroLog.Add("======================================");
                ficheroLog.Add("************ FIN PROCESO *************");
                ficheroLog.Add("======================================");

                BorrarDirectorio(medida_cuarto_horaria ? "inbox_cuartohoraria" : "inbox", "todos");


                // EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("ES02255021D");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                mes.SendMail("rsiope.gma@enel.com", p.GetValue("mail_cc",DateTime.Now, DateTime.Now), null, "TPLS Horarios Importados " + DateTime.Now.ToString("dd/MM/yyyy"),
                    System.Environment.NewLine +
                    System.Environment.NewLine +
                    "Ha finalizado el proceso de carga TPLS Horarios.", null);

            }
            catch(Exception e)
            {
                ficheroLog.AddError(e.Message);
            }
        }

        private void BorrarDirectorio(string parametro_ruta_archivos, string tipo_archivo)
        {
            string[] files = Directory.GetFiles(p.GetValue(parametro_ruta_archivos, DateTime.Now, DateTime.Now), "*P1*");
            for (int i = 0; i < files.Count(); i++)
            {
                FileInfo fichero = new FileInfo(files[i]);
                fichero.Delete();
            }

        }

        private bool ProcesaDirectorio(string parametro_ruta_archivos, string tipo_archivo)
        {
            bool hayError = false;
            bool firstOnly = true;
            string[] files = Directory.GetFiles(p.GetValue(parametro_ruta_archivos, DateTime.Now, DateTime.Now), "*P1*");
            { 
}
            try
            {


                for (int i = 0; i < files.Count(); i++)
                {
                    if (firstOnly)
                    {
                        Console.WriteLine("");
                        Console.WriteLine("Procesando " + tipo_archivo + " de un total de " + files.Count() + " archivos.");
                        ficheroLog.Add("Procesando " + tipo_archivo + " de un total de " + files.Count() + " archivos.");
                        firstOnly = false;
                    }

                    FileInfo fichero = new FileInfo(files[i]);
                    Console.CursorLeft = 0;
                    Console.Write("Procesando: " + fichero.Name + " " + (fichero.Length / 1024) + " KB");
                    ficheroLog.Add("Procesando: " + fichero.Name + " --> " + tipo_archivo + " " + (fichero.Length / 1024) + " KB");
                    hayError = LeerArchivo(fichero);

                    if (hayError)
                        fichero.MoveTo(p.GetValue("ruta_archivos_error", DateTime.Now, DateTime.Now) + "//" + fichero.Name);
                    else
                    {
                        GuardaArchivo(fichero, tipo_archivo);
                        // fichero.Delete();
                    }

                }
                // Finalmente pasamos a MySQL el resto de datos en memoria
                PasaMemoria_a_MySQL_Temp();
                return false;
            }catch(Exception e)
            {
                return true;
            }
        }

        private bool LeerArchivo(FileInfo file)
        {
            System.IO.StreamReader fileStream;
            string line;
            string[] c;
            int numLinea = 0;
            bool hayError = false;

            DateTime fecha = new DateTime();
            int horas = 0;
            int minutos = 0;
            int segundos = 0;


            fileStream = new System.IO.StreamReader(file.FullName, System.Text.Encoding.GetEncoding(1252));

            try
            {
                while ((line = fileStream.ReadLine()) != null)
                {
                    c = line.Split(';');
                    numLinea++;

                    fecha = new DateTime(Convert.ToInt32(c[2].Substring(0, 4)),
                        Convert.ToInt32(c[2].Substring(5, 2)),
                        Convert.ToInt32(c[2].Substring(8, 2)));
                    horas = Convert.ToInt32(c[2].Substring(11, 2));
                    minutos = Convert.ToInt32(c[2].Substring(14, 2));
                    segundos = Convert.ToInt32(c[2].Substring(17, 2));

                    PO1011 p = new PO1011();
                    p.cups22 = c[0];
                    p.tipo_medida = Convert.ToInt32(c[1]);
                    p.fecha_hora = fecha;
                    p.fecha_hora = p.fecha_hora.AddHours(horas);
                    p.fecha_hora = p.fecha_hora.AddMinutes(minutos);
                    p.fecha_hora = p.fecha_hora.AddSeconds(segundos);
                    p.estacion = Convert.ToInt32(c[3]);

                    if (c[4] != "")
                        p.ae = c[4].IndexOf(".") > 0 ? Convert.ToDouble(c[4].Substring(0, c[4].IndexOf("."))) : Convert.ToDouble(c[4]);
                    else
                        p.ae_nulo = true;
                    if (c[8] != "")
                        p.r[0] = c[8].IndexOf(".") > 0 ? Convert.ToDouble(c[8].Substring(0, c[8].IndexOf("."))) : Convert.ToDouble(c[8]);
                    else
                        p.r_nulo = true;

                    p.archivo = file.Name;

                    l_p.Add(p);

                    if (l_p.Count() > 1000000)
                        PasaMemoria_a_MySQL_Temp();
                }

                fileStream.Close();
                fileStream = null;
                return hayError;
            }
            catch (Exception e)
            {
                ficheroLog.AddError("ImportacionTPLs.LeerArchivo " + e.Message);
                return true;
            }
        }

        private void PasaMemoria_a_MySQL_Temp()
        {
            VuelcaMySQL(l_p);
            l_p.Clear();

        }

        private void VuelcaMySQL(List<PO1011> lc)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int numReg = 0;

            int x = 0;
            try
            {

                for (int i = 0; i < lc.Count(); i++)
                {
                    numReg++;
                    x++;

                    if (firstOnly)
                    {
                        sb.Append("replace into tpls (archivo,cups22,fecha,ae,r1,cups20,f_ult_mod) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(lc[i].archivo).Append("',");
                    sb.Append("'").Append(lc[i].cups22).Append("',");
                    sb.Append("'").Append(lc[i].fecha_hora.ToString("yyyy-MM-dd HH:mm:ss")).Append("',");

                    if (!lc[i].ae_nulo)
                        sb.Append(lc[i].ae.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    if (!lc[i].r_nulo)
                        sb.Append(lc[i].r[0].ToString().Replace(",", ".")).Append(", ");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append(lc[i].cups22.Substring(0, 20)).Append("',");
                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                    if (numReg == 5000)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.MED);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        numReg = 0;
                    }

                }

                if (numReg > 0)
                {
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                }




            }
            catch (Exception e)
            {
                ficheroLog.AddError("ImportacionTPLs.VuelcaMySQL " + e.Message);
            }
        }

        private void GuardaArchivo(FileInfo file, string origen)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            string strSql;
            try
            {
                strSql = "replace into tpls_archivos set "
                    + " archivo = '" + file.Name + "',"
                    + " peso = " + Convert.ToInt32(file.Length / 1024) + ","
                    + " origen = '" + origen + "',"
                    + " f_ult_mod = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                ficheroLog.AddError("ImportacionTPLs.GuardaArchivo: " + e.Message);
            }
        }

        private void SinValores_AE_R1()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            string strSql;
            try
            {
                ficheroLog.Add("SinValores_AE_R1");
                ficheroLog.Add("================");

                strSql = "replace into tpls_sin_ae_r1 SELECT * FROM tpls t WHERE t.ae IS NULL OR t.r1 IS null";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                ficheroLog.AddError("SinValores_AE_R1.GuardaArchivo: " + e.Message);
            }
        }

        private void Borrar_tabla(string tabla)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            string strSql;
            try
            {
                Console.WriteLine("Borrando datos en tabla " + tabla);
                ficheroLog.Add("Inicio borrado de datos de " + tabla);
                strSql = "delete from " + tabla;
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                ficheroLog.Add("Fin borrado de datos de " + tabla);
            }
            catch (Exception e)
            {
                ficheroLog.AddError("ImportacionTPLs.Borrar_tabla: " + e.Message);
            }
        }

        private void ANEXA_TPL_A_CCHODLAST()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            string strSql;
            try
            {
                ficheroLog.Add("Inicio anexión datos desde tpls a cchodlast");
                ficheroLog.Add("===========================================");
                strSql = "replace into " + prefijo_tabla + "cchodlast select archivo, cups22, fecha, ae, r1, cups20 from tpls where"
                    // + " cups22 IS NOT NULL and ae is not null and r1 is not null";
                    // Hay que incluir la reactiva aunque esté sin informar
                    + " cups22 IS NOT NULL and ae is not null";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                ficheroLog.Add("Fin anexión datos desde tpls a cchodlast");
            }
            catch (Exception e)
            {
                ficheroLog.AddError("ImportacionTPLs.ANEXA_TPL_A_CCHODLAST: " + e.Message);
            }
        }

        private void Volcado_de_cchodlast_A_med_cc_horaria_odhistorico_tmp(string cups22, PS ps)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime fecha = new DateTime();
            bool firstOnly = true;
            int p = 0;

            Dictionary<string, Dictionary<DateTime, EndesaEntity.medida.Med_cc_horaria_tmp>> d_cc = new Dictionary<string, Dictionary<DateTime, Med_cc_horaria_tmp>>();

            try
            {

                Console.CursorLeft = 0;
                Console.Write("Consultando datos para " + cups22);

                strSql = "select archivo, CUPS22, FECHA, AE, R1, CUPS20 CUPSREE from " + prefijo_tabla + "cchodlast where" +
                    " CUPS22 = '" + cups22 + "' order by FECHA";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fecha = Convert.ToDateTime(r["FECHA"]);
                    fecha = fecha.AddHours(-1);

                    p = fecha.Hour + 1;
                    if (fecha.Month == 10)
                        if (fecha.Date == UltimoDomingoOctubre(fecha) && fecha.Hour >= 1)
                            if (firstOnly)
                                firstOnly = false;
                            else
                                p++;


                    Dictionary<DateTime, EndesaEntity.medida.Med_cc_horaria_tmp> o;
                    if (!d_cc.TryGetValue(r["CUPS22"].ToString(), out o))
                    {
                        Med_cc_horaria_tmp c = new Med_cc_horaria_tmp();
                        c.cups22 = r["CUPS22"].ToString();
                        c.dia = fecha.Date;
                        c.a[p] = r["ae"].ToString();
                        c.r[p] = r["r1"].ToString();
                        c.archivo = r["archivo"].ToString();

                        o = new Dictionary<DateTime, Med_cc_horaria_tmp>();
                        o.Add(c.dia, c);
                        d_cc.Add(c.cups22, o);
                    }
                    else
                    {
                        Med_cc_horaria_tmp oo;
                        if (!o.TryGetValue(fecha.Date, out oo))
                        {
                            Med_cc_horaria_tmp c = new Med_cc_horaria_tmp();
                            c.cups22 = r["CUPS22"].ToString();
                            c.dia = fecha.Date;
                            c.a[p] = r["ae"].ToString();
                            c.r[p] = r["r1"].ToString();
                            c.archivo = r["archivo"].ToString();
                            o.Add(c.dia, c);
                        }
                        else
                        {
                            oo.a[p] = r["ae"].ToString();
                            oo.r[p] = r["r1"].ToString();
                        }
                    }

                }
                db.CloseConnection();

                VuelcaMySQL_CCHODLAST(d_cc, ps);

                //ficheroLog.Add("Fin CCHODLAST_A_CCHODLAST_HISTORICO");
                //ficheroLog.Add("===================================");
            }
            catch (Exception e)
            {
                ficheroLog.AddError("ImportacionTPLs.Volcado_de_cchodlast_A_med_cc_horaria_odhistorico_tmp --> CUPS " + cups22 + " " + e.Message);
            }
        }

        public void Volcado_de_med_cc_horaria_odhistorico_tmp_A_med_cc_horaria_odhistorico()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            servidores.MySQLDB db2;
            MySqlCommand command2;
            MySqlDataReader r2;
            string strSql = "";

            Dictionary<string, Dictionary<DateTime, EndesaEntity.medida.Med_cc_horaria_tmp>> d_cc = new Dictionary<string, Dictionary<DateTime, Med_cc_horaria_tmp>>();

            try
            {

                Console.WriteLine("Volcado de datos desde " + prefijo_tabla + "med_cc_horaria_odhistorico_tmp a "
                    + prefijo_tabla + "med_cc_horaria_odhistorico para días completos.");
                strSql = "REPLACE INTO " + prefijo_tabla + "med_cc_horaria_odhistorico"
                    + " SELECT idmp, fecha, sumdailya, sumdailyr,"
                    + " a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25,"
                    + " r1, r2, r3, r4, r5, r6, r7, r8, r9, r10, r11, r12, r13, r14, r15, r16, r17, r18, r19, r20, r21, r22, r23, r24, r25, archivo"
                    + " FROM " + prefijo_tabla + "med_cc_horaria_odhistorico_tmp t WHERE t.completo = 'S'";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // OK

                Console.WriteLine("Borrado de datos en " + prefijo_tabla + "med_cc_horaria_odhistorico_tmp para días completos.");
                strSql = "delete from " + prefijo_tabla + "med_cc_horaria_odhistorico_tmp WHERE completo = 'S'";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Volcado de datos desde " + prefijo_tabla + "med_cc_horaria_odhistorico_tmp a "
                    + prefijo_tabla + "med_cc_horaria_odhistorico idmp y fecha que no existe.");

                strSql = "replace into " + prefijo_tabla + "med_cc_horaria_odhistorico_tmp2 SELECT t.idmp,t.fecha,'N',t.sumdailya,t.sumdailyr,"
                    + " t.a1,t.a2,t.a3,t.a4,t.a5,t.a6,t.a7,t.a8,t.a9,t.a10,t.a11,t.a12,t.a13,t.a14,t.a15,t.a16,t.a17,t.a18,t.a19,t.a20,t.a21,t.a22,t.a23,t.a24,t.a25,"
                    + " t.r1,t.r2,t.r3,t.r4,t.r5,t.r6,t.r7,t.r8,t.r9,t.r10,t.r11,t.r12,t.r13,t.r14,t.r15,t.r16,t.r17,t.r18,t.r19,t.r20,t.r21,t.r22,t.r23,t.r24,t.r25,t.archivo"
                    + " FROM " + prefijo_tabla + "med_cc_horaria_odhistorico_tmp t"
                    + " LEFT OUTER JOIN " + prefijo_tabla + "med_cc_horaria_odhistorico m ON"
                    + " m.IDMP = t.idmp AND"
                    + " m.DIA = t.fecha"
                    + " WHERE m.IDMP IS NULL;";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "replace into " + prefijo_tabla + "med_cc_horaria_odhistorico SELECT t.idmp,t.fecha,t.sumdailya,t.sumdailyr,"
                    + " t.a1,t.a2,t.a3,t.a4,t.a5,t.a6,t.a7,t.a8,t.a9,t.a10,t.a11,t.a12,t.a13,t.a14,t.a15,t.a16,t.a17,t.a18,t.a19,t.a20,t.a21,t.a22,t.a23,t.a24,t.a25,"
                    + " t.r1,t.r2,t.r3,t.r4,t.r5,t.r6,t.r7,t.r8,t.r9,t.r10,t.r11,t.r12,t.r13,t.r14,t.r15,t.r16,t.r17,t.r18,t.r19,t.r20,t.r21,t.r22,t.r23,t.r24,t.r25,t.archivo"
                    + " FROM " + prefijo_tabla + "med_cc_horaria_odhistorico_tmp2 t"
                    + " LEFT OUTER JOIN " + prefijo_tabla + "med_cc_horaria_odhistorico m ON"
                    + " m.IDMP = t.idmp AND"
                    + " m.DIA = t.fecha"
                    + " WHERE m.IDMP IS NULL;";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "DELETE t.* FROM " + prefijo_tabla + "med_cc_horaria_odhistorico_tmp t"
                    + " LEFT OUTER JOIN " + prefijo_tabla + "med_cc_horaria_odhistorico_tmp2 m ON"
                    + " m.IDMP = t.idmp AND m.fecha = t.fecha"
                    + " WHERE m.IDMP IS NULL; ";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "delete from " + prefijo_tabla + "med_cc_horaria_odhistorico_tmp2";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "select * from " + prefijo_tabla + "med_cc_horaria_odhistorico_tmp";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    strSql = "Select "
                    + " IDMP,DIA,SumDailyA,SumDailyR,"
                    + " a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,r19,a20,a21,a22,a23,a24,a25,"
                    + " r1,r2,r3,r4,r5,r6,r7,r8,r9,r10,r11,r12,r13,r14,r15,r16,r17,r18,r19,r20,r21,r22,r23,r24,r25,archivo"
                        + " from " + prefijo_tabla + "med_cc_horaria_odhistorico where"
                        + " IDMP = " + r["idmp"].ToString() + " and"
                        + " DIA = '" + Convert.ToDateTime(r["fecha"]).ToString("yyyy-MM-dd") + "'";
                    db2 = new MySQLDB(MySQLDB.Esquemas.MED);
                    command2 = new MySqlCommand(strSql, db2.con);
                    r2 = command.ExecuteReader();
                    while (r2.Read())
                    {
                        strSql = "update " + prefijo_tabla + "med_cc_horaria_odhistorico set"
                            + " Archivo = '" + r["archivo"].ToString() + "'";

                        for (int i = 1; i <= 25; i++)
                            if (r["a" + i] != System.DBNull.Value)
                                strSql += ",A" + i + " = " + r2["a" + i].ToString().Replace(",", ".");

                        for (int i = 1; i <= 25; i++)
                            if (r["r" + i] != System.DBNull.Value)
                                strSql += ",R" + i + " = " + r2["r" + i].ToString().Replace(",", ".");

                        strSql += " WHERE IDMP = " + r["idmp"].ToString() + " and"
                             + " DIA = '" + Convert.ToDateTime(r["fecha"]).ToString("yyyy-MM-dd") + "'";
                    }
                    db2.CloseConnection();
                }
                db.CloseConnection();



            }
            catch (Exception e)
            {
                ficheroLog.AddError("ImportacionTPLs.Volcado_de_med_cc_horaria_odhistorico_tmp_A_med_cc_horaria_odhistorico " + strSql + " " + e.Message);
            }
        }

        private void VuelcaMySQL_CCHODLAST(Dictionary<string, Dictionary<DateTime, EndesaEntity.medida.Med_cc_horaria_tmp>> d, PS ps)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int numReg = 0;
            double totala = 0;
            double totalr = 0;
            int totalPeriodos = 0;

            int x = 0;
            try
            {

                foreach (KeyValuePair<string, Dictionary<DateTime, EndesaEntity.medida.Med_cc_horaria_tmp>> p in d)
                {
                    numReg++;
                    x++;

                    foreach (KeyValuePair<DateTime, EndesaEntity.medida.Med_cc_horaria_tmp> pp in p.Value)
                    {

                        totalPeriodos = 0;

                        if (firstOnly)
                        {
                            sb.Append("replace into " + prefijo_tabla + "med_cc_horaria_odhistorico_tmp (idmp,fecha,completo,sumdailya,sumdailyr");
                            for (int i = 1; i <= 25; i++)
                                sb.Append(",a" + i);
                            for (int i = 1; i <= 25; i++)
                                sb.Append(",r" + i);

                            sb.Append(",archivo) values ");
                            firstOnly = false;
                        }

                        sb.Append("(").Append(ps.Get_IDMP(p.Key)).Append(",");
                        sb.Append("'").Append(pp.Value.dia.ToString("yyyy-MM-dd")).Append("',");

                        for (int i = 1; i <= 25; i++)
                            if (pp.Value.a[i] != null)
                                totalPeriodos++;

                        sb.Append(totalPeriodos == NumPeriodosHorarios(pp.Value.dia) ? "'S'," : "'N',");

                        totala = 0;
                        totalr = 0;
                        for (int i = 1; i <= 25; i++)
                        {
                            totala += Convert.ToDouble(pp.Value.a[i]);
                            totalr += Convert.ToDouble(pp.Value.r[i]);
                        }

                        sb.Append(totala).Append(",");
                        sb.Append(totalr);

                        for (int i = 1; i <= 25; i++)
                            sb.Append(",").Append(pp.Value.a[i] != null ? pp.Value.a[i].Replace(",", ".") : "null");


                        for (int i = 1; i <= 25; i++)
                            sb.Append(",").Append(pp.Value.r[i] != null ? pp.Value.r[i].Replace(",", ".") : "null");

                        sb.Append(",'").Append(pp.Value.archivo).Append("'),");


                        if (numReg == 250)
                        {
                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.MED);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            numReg = 0;
                        }

                    }
                }

                if (numReg > 0)
                {
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                }

            }
            catch (Exception e)
            {
                ficheroLog.AddError("ImportacionTPLs.VuelcaMySQL " + e.Message);
            }
        }

        private DateTime UltimoDomingoMarzo(DateTime f)
        {
            return new DateTime(f.Year, 3, 31).AddDays(-(new DateTime(f.Year, 3, 31).DayOfWeek - DayOfWeek.Sunday));
        }

        private DateTime UltimoDomingoOctubre(DateTime f)
        {
            return new DateTime(f.Year, 10, 31).AddDays(-(new DateTime(f.Year, 10, 31).DayOfWeek - DayOfWeek.Sunday));
        }

        private int NumPeriodosCuartoHorarios(DateTime f)
        {
            return (f.Month == 3 ?
                    f.Equals(UltimoDomingoMarzo(f)) ? 92 : 96 :
                    f.Month == 10 ?
                    f.Equals(UltimoDomingoOctubre(f)) ? 100 : 96 : 96);
        }

        private int NumPeriodosHorarios(DateTime f)
        {
            return (f.Month == 3 ?
                    f.Equals(UltimoDomingoMarzo(f)) ? 23 : 24 :
                    f.Month == 10 ?
                    f.Equals(UltimoDomingoOctubre(f)) ? 25 : 24 : 24);
        }

        private void ArreglaUltimoDomingoMarzo(bool medida_cuartohoraria)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader rs;
            servidores.MySQLDB db2;
            MySqlCommand command2;
            MySqlDataReader rs2;
            servidores.MySQLDB db3;
            MySqlCommand command3;            
            string strSql;


            try
            {
                strSql = "select YEAR(FECHA) as year from " + prefijo_tabla + "med_cc_horaria_odhistorico_temporal group by year(FECHA)";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                rs = command.ExecuteReader();
                while (rs.Read())
                {
                    DateTime fecha = UltimoDomingoMarzo(new DateTime(Convert.ToInt32(rs["year"]), 3, 1));

                    strSql = "select CUPS20, CUPS22, FECHA, Archivo, A3, R3 from " + prefijo_tabla + "med_cc_horaria_odhistorico_temporal where"
                        + " FECHA = '" + fecha.ToString("yyyy-MM-dd") + "' and"
                        + " A3 is not null";
                    db2 = new MySQLDB(MySQLDB.Esquemas.MED);
                    command2 = new MySqlCommand(strSql, db2.con);
                    rs2 = command.ExecuteReader();
                    while (rs2.Read())
                    {
                        strSql = "replace into " + prefijo_tabla + "med_cc_horaria_odhistorico_temporal (CUPS20, CUPS22, FECHA, Archivo, A2, R2) values"
                            + " ('" + rs2["cups20"].ToString() + "',"
                            + " '" + rs2["cups22"].ToString() + "',"
                            + "'" + Convert.ToDateTime(rs2["fecha"]).ToString("yyyy-MM-dd") + "',"
                            + " '" + rs2["archivo"].ToString() + "',"
                            + rs2["A3"].ToString().Replace(", ", ".") + " ,"
                            + rs2["R3"].ToString().Replace(", ", ".") + ")";
                        db3 = new MySQLDB(MySQLDB.Esquemas.MED);
                        command3 = new MySqlCommand(strSql, db3.con);
                        command3.ExecuteNonQuery();
                        db3.CloseConnection();

                        strSql = "delete from " + prefijo_tabla + "med_cc_horaria_odhistorico_temporal where"
                            + " CUPS20 = '" + rs2["cups20"].ToString() + "' and"
                            + " CUPS22 = '" + rs2["cups22"].ToString() + "' and"
                            + " FECHA = '" + Convert.ToDateTime(rs2["fecha"]).ToString("yyyy-MM-dd") + "' and"
                            + " Archivo = '" + rs2["archivo"].ToString() + "' and"
                            + " A3 = " + rs2["A3"].ToString().Replace(",", ".");
                        db3 = new MySQLDB(MySQLDB.Esquemas.MED);
                        command3 = new MySqlCommand(strSql, db3.con);
                        command3.ExecuteNonQuery();
                        db3.CloseConnection();
                    }
                    db2.CloseConnection();

                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                ficheroLog.AddError("ImportacionTPLs.ArreglaUltimoDomingoMarzo: " + e.Message);
            }
        }
    }
}
