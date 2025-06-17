using EndesaBusiness.servidores;
using EndesaEntity.global;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion
{
    public class GestionPropiaATR
    {
        Dictionary<string, EndesaEntity.contratacion.GestionPropia> dic = new Dictionary<string, EndesaEntity.contratacion.GestionPropia>();
        List<EndesaEntity.PuntoSuministro> lc = new List<EndesaEntity.PuntoSuministro>();
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_GestionPropiaATR");
        utilidades.Param param;
        DateTime ultimoDiaHabil;
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;

        public GestionPropiaATR()
        {
            param = new utilidades.Param("gestionpropiaatr_param", servidores.MySQLDB.Esquemas.CON);
            ultimoDiaHabil = utilidades.Fichero.UltimoDiaHabil();
            ss_pp = new utilidades.Seguimiento_Procesos();
            
        }

        public void DescargaAtrs()
        {
            string strSql = "";
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            DateTime fecha_GestionPropiaATR = new DateTime();

            try
            {
                // Comprobamos fecha Proceso GestionPropiaATR
                strSql = "select fecha from ps_fechas_procesos where"
                    + " proceso = 'gestionATR'";
                db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    fecha_GestionPropiaATR = Convert.ToDateTime(reader["fecha"]);
                }
                db.CloseConnection();


                if (fecha_GestionPropiaATR.Date < DateTime.Now.Date)
                {
                    ss_pp.Update_Fecha_Inicio("Contratación", "PS_AT", "Ejecución Extractor Gestion Propia ATR");
                    ficheroLog.Add("Ejecutando extractor: " + param.GetValue("atr", DateTime.Now, DateTime.Now));
                    utilidades.Fichero.EjecutaComando(param.GetValue("atr", DateTime.Now, DateTime.Now));
                    ss_pp.Update_Fecha_Fin("Contratación", "PS_AT", "Ejecución Extractor Gestion Propia ATR");
                }
            }catch(Exception ex)
            {
                ss_pp.Update_Comentario("Contratación", "PS_AT", "Ejecución Extractor Gestion Propia ATR", ex.Message);
                ficheroLog.AddError("DescargaAtrs: " + ex.Message);
            }

            

        }

        public void DescargaPuntosGestionPropiaATR()
        {
            try
            {
                if (DescargarGestionPropiaATR())
                {
                    ficheroLog.Add("Ejecutando extractor: " + param.GetValue("gestionpropiaatr", DateTime.Now, DateTime.Now));
                    utilidades.Fichero.EjecutaComando(param.GetValue("gestionpropiaatr", DateTime.Now, DateTime.Now), ultimoDiaHabil.ToString("yyMMdd"));
                }


            }
            catch (Exception e)
            {
                ficheroLog.AddError("DescargaPuntosGestionPropiaATR --> " + e.Message);
            }


        }


        public void LanzaGestionPropiaATR()
        {
            string strSql = "";
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            DateTime fecha_GestionPropiaATR = new DateTime();
            DateTime hoy = new DateTime();

            // Comprobamos fecha Proceso GestionPropiaATR
            strSql = "select fecha from ps_fechas_procesos where"
                + " proceso = 'gestionATR'";
            db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            reader = command.ExecuteReader();
            if (reader.Read())
            {
                fecha_GestionPropiaATR = Convert.ToDateTime(reader["fecha"]);
            }
            db.CloseConnection();

            if (fecha_GestionPropiaATR.Date < DateTime.Now.Date)
            {
                ss_pp.Update_Fecha_Inicio("Contratación", "PS_AT", "Importación Gestion Propia ATR");
                hoy = DateTime.Now;
                if (hoy >= new DateTime(2018, 11, 26))
                    CargaPuntosGestionPropiaATR_NuevaExtraccion();
                else
                    CargaPuntosGestionPropiaATR();

                ss_pp.Update_Fecha_Fin("Contratación", "PS_AT", "Importación Gestion Propia ATR");
            }

        }

        private void CargaPuntosGestionPropiaATR()
        {
            FileInfo f;
            string[] listaArchivos;
            string rutaOrigen = "";
            string nombreArchivo = "";
            string line = "";
            string[] c;

            int i = 0;
            int j = 0;
            int numreg = 0;
            StringBuilder sb = new StringBuilder();
            office.MailCompose mc;            
            utilidades.Global g = new utilidades.Global();
            double filesize;

            try
            {

                rutaOrigen = param.GetValue("carpeta_entrada", DateTime.Now, DateTime.Now);
                nombreArchivo = param.GetValue("prefijo_archivo1", DateTime.Now, DateTime.Now) + ultimoDiaHabil.ToString("yyMMdd") + ".txt";
                listaArchivos = Directory.GetFiles(rutaOrigen, nombreArchivo);

                f = new FileInfo(rutaOrigen + "\\" + nombreArchivo);
                if (f.Exists && f.Length > 10000)
                {

                    ficheroLog.Add("Cargando " + f.FullName);

                    System.IO.StreamReader archivo = new System.IO.StreamReader(f.FullName, System.Text.Encoding.GetEncoding(1252));
                    while ((line = archivo.ReadLine()) != null)
                    {
                        i = 0;
                        c = line.Split(';');
                        numreg++;
                        EndesaEntity.contratacion.GestionPropia atr = new EndesaEntity.contratacion.GestionPropia();
                        atr.estadoCont = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.cups = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.cliente = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.faltacont = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.fbajacont = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.fprevbajacont = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.tarifa = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.tipocli = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.provincia = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.minicipio = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.lineanegocio = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.sva = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.segmentomercado = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.nif = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;

                        if (atr.cups.Length >= 13 && (atr.cups.Substring(0, 3) != "XXX" && atr.lineanegocio == "001" && atr.sva == "E" && atr.segmentomercado == "N"))
                        {

                            EndesaEntity.contratacion.GestionPropia a;
                            if (!dic.TryGetValue(atr.cups, out a))
                            {
                                dic.Add(atr.cups, atr);
                                lc.Add(new EndesaEntity.PuntoSuministro() { cups13 = atr.cups, id = j });
                                j++;
                            }

                        }
                    }
                    filesize = (f.Length / 1024);
                    SaveMD5(f.Name, "GestionPropiaATR", numreg, dic.Count(), utilidades.Fichero.checkMD5(f.FullName).ToString(), filesize);

                    archivo.Close();

                    if (dic.Count() > Convert.ToInt32(param.GetValue("puntos_min_gestionpropia", DateTime.Now, DateTime.Now)))
                    {
                        CompletaCUPS22();
                        GuardaGestionPropiaBBDD();
                        ActualizaDistribuidora();
                        BorraPuntosCEFACO();
                        ActualizaFechaProceso(utilidades.Fichero.SiguienteDiaHabil());
                        ActualizaProcesoAccess(utilidades.Fichero.SiguienteDiaHabil());
                        ActualizaFechaProceso_OK(DateTime.Now);
                        // Informamos que el proceso ha terminado satisfactoriamente
                        mc = new office.MailCompose(servidores.MySQLDB.Esquemas.CON, "gestionpropiaatr_mail", "GestionPropiaATR");
                        mc.Send();
                    }
                    else
                    {
                        ActualizaFechaProceso(utilidades.Fichero.SiguienteDiaHabil());
                        // Quitado el 20181115 porque se ha quitado de la cola de procesos Access
                        //ActualizaProcesoAccess(utilidades.Fichero.SiguienteDiaHabil());

                        // Informamos que hay un error en el proceso.
                        mc = new office.MailCompose(servidores.MySQLDB.Esquemas.CON, "gestionpropiaatr_mail", "ErrorGestionPropiaATR");
                        mc.body = mc.body.Replace("@2", UltimaEjecucionOK().ToString("dd/MM/yyyy"));
                        mc.Send();
                    }

                    f.Delete();
                }
            }
            catch (Exception e)
            {
                ficheroLog.AddError(e.Message);
            }

        }
        private void CargaPuntosGestionPropiaATR_NuevaExtraccion()
        {
            FileInfo f;
            string[] listaArchivos;
            string rutaOrigen = "";
            string nombreArchivo = "";
            string line = "";
            string[] c;

            int i = 0;
            int j = 0;
            int numreg = 0;
            StringBuilder sb = new StringBuilder();
            office.MailCompose mc;
            utilidades.Global g = new utilidades.Global();
            double filesize;

            try
            {

                rutaOrigen = param.GetValue("carpeta_entrada", DateTime.Now, DateTime.Now);
                nombreArchivo = param.GetValue("prefijo_archivo1", DateTime.Now, DateTime.Now) + ultimoDiaHabil.ToString("yyMMdd") + ".txt";
                listaArchivos = Directory.GetFiles(rutaOrigen, nombreArchivo);

                f = new FileInfo(rutaOrigen + "\\" + nombreArchivo);
                if (f.Exists)
                {

                    ficheroLog.Add("Cargando " + f.FullName);

                    System.IO.StreamReader archivo = new System.IO.StreamReader(f.FullName, System.Text.Encoding.GetEncoding(1252));
                    while ((line = archivo.ReadLine()) != null)
                    {
                        i = 0;
                        c = line.Split(';');
                        numreg++;
                        EndesaEntity.contratacion.GestionPropia atr = new EndesaEntity.contratacion.GestionPropia();
                        atr.estadoCont = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.cups = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.cliente = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.faltacont = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.fbajacont = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.fprevbajacont = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.tarifa = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.tipocli = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.provincia = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.minicipio = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.lineanegocio = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.sva = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.segmentomercado = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;
                        atr.nif = utilidades.FuncionesTexto.ArreglaAcentos(c[i]); i++;

                        EndesaEntity.contratacion.GestionPropia a;
                        if (!dic.TryGetValue(atr.cups, out a))
                        {
                            dic.Add(atr.cups, atr);
                            lc.Add(new EndesaEntity.PuntoSuministro() { cups13 = atr.cups, id = j });
                            j++;
                        }


                    }

                    filesize = (f.Length / 1024);
                    SaveMD5(f.Name, "GestionPropiaATR", numreg, dic.Count(), utilidades.Fichero.checkMD5(f.FullName).ToString(), filesize);

                    archivo.Close();

                    if (dic.Count() > Convert.ToInt32(param.GetValue("puntos_min_gestionpropia", DateTime.Now, DateTime.Now)))
                    {
                        CompletaCUPS22();
                        GuardaGestionPropiaBBDD();
                        ActualizaDistribuidora();
                        BorraPuntosCEFACO();
                        ActualizaFechaProceso(utilidades.Fichero.SiguienteDiaHabil());
                        
                        ActualizaFechaProceso_OK(DateTime.Now);
                        
                        mc = new office.MailCompose(servidores.MySQLDB.Esquemas.CON, "gestionpropiaatr_mail", "GestionPropiaATR");
                        mc.Send();
                    }
                    else
                    {
                        ActualizaFechaProceso(utilidades.Fichero.SiguienteDiaHabil());
                        //ActualizaProcesoAccess(utilidades.Fichero.SiguienteDiaHabil());

                        // Informamos que hay un error en el proceso.
                        mc = new office.MailCompose(servidores.MySQLDB.Esquemas.CON, "gestionpropiaatr_mail", "ErrorGestionPropiaATR");
                        mc.body = mc.body.Replace("@2", UltimaEjecucionOK().ToString("dd/MM/yyyy"));
                        mc.Send();
                    }

                    f.Delete();
                }
            }
            catch (Exception e)
            {
                ficheroLog.AddError("CargaPuntosGestionPropiaATR_NuevaExtraccion: " + e.Message);
            }

        }

        private DateTime UltimaEjecucionOK()
        {
            string strSql = "";
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            DateTime fecha_GestionPropiaATR = new DateTime();

            // Comprobamos la última ejeción OK de GestionPropiaATR
            strSql = "select fecha from ps_fechas_procesos where"
                + " proceso = 'gestionATR_OK'";
            db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            reader = command.ExecuteReader();
            if (reader.Read())
            {
                fecha_GestionPropiaATR = Convert.ToDateTime(reader["fecha"]);
            }
            db.CloseConnection();

            return fecha_GestionPropiaATR;
        }


        private void ActualizaDistribuidora()
        {
            string strSql = "";
            servidores.MySQLDB db;
            MySqlCommand command;

            strSql = "update gestionpropiaatr atr"
                + " inner join paramDistribuidoras d on"
                + " d.CDISTRIB = substr(atr.CUPS, 1, 3)"
                + " set atr.distribuidora = d.Distribuidora";

            db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        private void BorraPuntosCEFACO()
        {
            string strSql = "";
            servidores.MySQLDB db;
            MySqlCommand command;

            strSql = "DELETE FROM gestionpropiaatr WHERE segmentoMercado <> 'N'";
            db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }


        private void ActualizaFechaProceso_OK(DateTime d)
        {
            string strSql = "";
            servidores.MySQLDB db;
            MySqlCommand command;

            ficheroLog.Add("Actualizamos la fecha del proceso gestionATR_OK de la tabla ps_fechas_procesos con valor " + d.ToString("yyyy-MM-dd"));
            strSql = "update ps_fechas_procesos"
            + " set fecha = '" + d.ToString("yyyy-MM-dd") + "'"
            + " where proceso = 'gestionATR_OK'";

            db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        private void ActualizaFechaProceso(DateTime d)
        {
            string strSql = "";
            servidores.MySQLDB db;
            MySqlCommand command;

            if (param.GetValue("habilita_fecha_proceso_mysql", DateTime.Now, DateTime.Now) == "S")
            {

                ficheroLog.Add("Actualizamos la fecha del proceso gestionATR de la tabla ps_fechas_procesos con valor " + d.ToString("yyyy-MM-dd"));
                strSql = "update ps_fechas_procesos"
                + " set fecha = '" + d.ToString("yyyy-MM-dd") + "'"
                + " where proceso = 'gestionATR'";

                db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }

        }

        private void ActualizaProcesoAccess(DateTime ini)
        {
            string strSql = "";
            servidores.AccessDB ac;
            string rutaAccessProcesos;

            try
            {
                if (param.GetValue("habilita_fecha_proceso_access", DateTime.Now, DateTime.Now) == "S")
                {
                    ficheroLog.Add("Actualizamos la tabla #_paramfechasProcesos del Access gestionPropiaATR.");
                    strSql = "UPDATE [#_paramfechasProcesos] set [ejecucionOK] = #" + utilidades.Fichero.SiguienteDiaHabil().ToString("MM/dd/yyyy") + "#";
                    rutaAccessProcesos = param.GetValue("Access", DateTime.Now, DateTime.Now);
                    ac = new servidores.AccessDB(rutaAccessProcesos);
                    OleDbCommand cmd = new OleDbCommand(strSql, ac.con);
                    cmd.ExecuteNonQuery();
                    ac.CloseConnection();
                }

            }
            catch (Exception e)
            {
                ficheroLog.AddError("ActualizaProcesoAccess " + e.Message);
            }


        }


        private void CompletaCUPS22()
        {
            ficheroLog.Add("Completamos CUPS22");
            cups.PuntosSuministro ps = new cups.PuntosSuministro();
            ps.CompletaCups22(lc);
            foreach (KeyValuePair<string, EndesaEntity.contratacion.GestionPropia> p in dic)
            {
                p.Value.cups22 = lc.Find(z => z.cups13 == p.Value.cups).cups20;
            }
        }

        private bool DescargarGestionPropiaATR()
        {

            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql = "";
            bool descargar = true;


            try
            {
                strSql = "select max(processdatetime) processdatetime from gestionpropiaatr_files where "
                    + " valid_numlines > " + param.GetValue("puntos_min_gestionpropia", DateTime.Now, DateTime.Now);


                db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    descargar = ultimoDiaHabil.Date >= Convert.ToDateTime(reader["processdatetime"]).Date;
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                ficheroLog.AddError("DescargarGestionPropiaATR --> " + e.Message);
            }

            return descargar;
        }


        private void GuardaGestionPropiaBBDD()
        {
            string strSql = "";
            string clave;
            bool firstOnly = true;
            int j = 0;
            int i = 0;
            int numreg = 0;
            StringBuilder sb = new StringBuilder();
            servidores.MySQLDB db;
            MySqlCommand command;

            try
            {
                ficheroLog.Add("Pasamos tabla gestionpropiaatr a historico");

                db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                strSql = "replace into gestionpropiaatr_hist select ps.*,"
                    + " '" + ultimoDiaHabil.ToString("yyyy-MM-dd") + "'"
                    + " from gestionpropiaatr ps";
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                ficheroLog.Add("Borramos tabla gestionpropiaatr ");
                db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand("delete from gestionpropiaatr", db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                foreach (KeyValuePair<string, EndesaEntity.contratacion.GestionPropia> p in dic)
                {
                    clave = p.Key;
                    j++;
                    numreg++;
                    i = 0;
                    if (firstOnly)
                    {
                        sb.Append("replace into gestionpropiaatr (estadoCont,CUPS,cliente,fAltaCont,fBajaCont,fPrevBajaCont,TARIFA,tipoCli,");
                        sb.Append("PROVINCIA,MUNICIPIO,lineaNegocio,SVA,segmentoMercado,NIF,CUPS22) values");
                        firstOnly = false;
                    }
                    i = 0;

                    sb.Append(" ('").Append(p.Value.estadoCont).Append("',"); i++;    // estadoCont
                    sb.Append("'").Append(p.Value.cups).Append("',"); i++;            // CUPS                    
                    sb.Append("'").Append(p.Value.cliente).Append("',"); i++;         // cliente
                    sb.Append("'").Append(p.Value.faltacont).Append("',"); i++;       // fAltaCont
                    sb.Append("'").Append(p.Value.fbajacont).Append("',"); i++;       // fBajaCont
                    sb.Append("'").Append(p.Value.fprevbajacont).Append("',"); i++;   // fPrevBajaCont
                    sb.Append("'").Append(p.Value.tarifa).Append("',"); i++;          // TARIFA
                    sb.Append("'").Append(p.Value.tipocli).Append("',"); i++;         // tipoCli
                    sb.Append("'").Append(p.Value.provincia).Append("',"); i++;       // PROVINCIA
                    sb.Append("'").Append(p.Value.minicipio).Append("',"); i++;       // MUNICIPIO
                    sb.Append("'").Append(p.Value.lineanegocio).Append("',"); i++;    // lineaNegocio
                    sb.Append("'").Append(p.Value.sva).Append("',"); i++;             // SVA
                    sb.Append("'").Append(p.Value.segmentomercado).Append("',"); i++; // segmentoMercado
                    sb.Append("'").Append(p.Value.nif).Append("',"); i++;             // NIF                    
                    sb.Append("'").Append(p.Value.cups22).Append("'),"); i++;         // CUPS22 

                    if (numreg == 250)
                    {
                        Console.WriteLine("Leyendo " + j + " lineas...");
                        firstOnly = true;
                        db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        numreg = 0;
                    }

                }


                if (numreg > 0)
                {
                    Console.WriteLine("Leyendo " + j + " lineas...");
                    firstOnly = true;
                    db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    numreg = 0;
                }
            }
            catch (Exception e)
            {
                ficheroLog.AddError(e.Message);
            }


        }

        private void SaveMD5(string file, string description, Int32 numLines, Int32 valid_numLines, string md5, double filesize)
        {

            servidores.MySQLDB db;
            MySqlCommand command;
            string strSql;
            try
            {
                db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                strSql = "REPLACE into gestionpropiaatr_files SET"
                 + " filename = '" + file + "',"
                 + " description = '" + description + "',"
                 + " processdatetime = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                 + " numlines = " + numLines + ","
                 + " valid_numlines = " + valid_numLines + ","
                 + " md5 = '" + md5 + "',"
                 + " filesize = " + filesize + ";";
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("SaveMD5 " + e.Message);
            }
        }

    }
}
