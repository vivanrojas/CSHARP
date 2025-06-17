using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace EndesaBusiness.contratacion
{
    public class SolicitudesATR : EndesaEntity.contratacion.SolATRMT
    {

        Dictionary<string, EndesaEntity.contratacion.SolATRMT> dic { get; set; }

        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "SolicitudesATR");
        utilidades.Param param;
        DateTime ultimoDiaHabil;
        List<EndesaEntity.contratacion.SolATRMT> lista;
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;

        public SolicitudesATR()
        {
            param = new utilidades.Param("solatrmt_param", servidores.MySQLDB.Esquemas.CON);
            ultimoDiaHabil = utilidades.Fichero.UltimoDiaHabil();
            ficheroLog.Add("Ultimo Día Hábil --> " + ultimoDiaHabil.ToString("yyMMdd"));
            lista = new List<EndesaEntity.contratacion.SolATRMT>();
            ss_pp = new utilidades.Seguimiento_Procesos();
        }

        public SolicitudesATR(List<string> lista_cups_20, 
            int empresa, string estadoSolATR, string tipoSolATR)
        {
            dic = Carga(lista_cups_20, empresa, estadoSolATR, tipoSolATR);
        }

        private Dictionary<string, EndesaEntity.contratacion.SolATRMT> Carga(List<string> lista_cups_20, 
            int empresa, string estadoSolATR, string tipoSolATR)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            Dictionary<string, EndesaEntity.contratacion.SolATRMT> d =
                new Dictionary<string, EndesaEntity.contratacion.SolATRMT>();

            try
            {
                strSql = "SELECT CUPS_EXT, codSolATR, fAcepRech"
                    + " FROM SOLATRMT"
                    + " where CUPS_EXT in ('" + lista_cups_20[0] + "'";

                for (int i = 1; i < lista_cups_20.Count; i++)
                    strSql += " ,'" + lista_cups_20[i] + "'";

                strSql += ") and EMP_TIT = " + empresa
                    + " AND estadoSolATR = '" + estadoSolATR + "'"
                    + " AND tipoSolATR = '" + tipoSolATR + "'"
                    + " ORDER BY CUPS_EXT, fAcepRech DESC";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.SolATRMT c = new EndesaEntity.contratacion.SolATRMT();
                    c.cups_ext = r["CUPS_EXT"].ToString();
                    c.fAcepRech = Convert.ToInt32(r["fAcepRech"]);

                    EndesaEntity.contratacion.SolATRMT o;
                    if (!d.TryGetValue(c.cups_ext, out o))
                        d.Add(c.cups_ext, c);
                }
                db.CloseConnection();

                return d;
            }catch(Exception e)
            {
                return null;
            }
        }


        public List<EndesaEntity.contratacion.SolATRMT> Carga_SOLATRMT_MAX(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            List<EndesaEntity.contratacion.SolATRMT> lista =
                new List<EndesaEntity.contratacion.SolATRMT>();

            try
            {
                strSql = "SELECT sol.EMP_TIT, sol.estadoSolATR, sol.codSolATR, sol.tipoSolATR, sol.fRecepcion,"
                   + " sol.fAcepRech, sol.fRechazo, sol.fEnvioATR, sol.fEnvioDoc, sol.fPrevAlta, sol.fActivacion,"
                   + " sol.CCOUNIPS, sol.CUPS_EXT, sol.TIPO_CLI, sol.CIF, sol.cliente, sol.SEG_MER, sol.LINEA_N,"
                   + " sol.DIREC_PS, sol.CDISTRIB, sol.CONTR_ATR,"
                   + " sol.COMENT_1_SOLICITUD, sol.COMENT_2_SOLICITUD, sol.COMENT_3_SOLICITUD, sol.CHECK_ANULAC,"
                   + " sol.MOT_BAJA, sol.POTENCIA1, sol.POTENCIA2, sol.POTENCIA3, sol.POTENCIA4, sol.POTENCIA5,"
                   + " sol.POTENCIA6, sol.TARIFA, sol.TENSION, sol.CCONTRPS, sol.VER_CONTR_PS, sol.USO_CONTR,"
                   + " sol.GESTOR_SCE, cartera.territorio, cartera.zona,"
                   + " cartera.responsable_territorial, rt.email_gestor"
                   + " FROM SOLATRMT_MAX sol"
                   + " LEFT OUTER JOIN salesforce_cartera cartera ON"
                   + " cartera.nif = sol.CIF"
                   + " LEFT OUTER JOIN salesforce_cartera_rt rt ON"
                   + " rt.responsable_territorial = cartera.responsable_territorial"
                   + " WHERE sol.fRecepcion >= " + fd.ToString("yyyyMMdd") + " and"
                   + " sol.fRecepcion <= " + fh.ToString("yyyyMMdd");
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.SolATRMT c =
                        new EndesaEntity.contratacion.SolATRMT();

                }
                db.CloseConnection();
                return lista;

            }
            catch(MySqlException e)
            {
                return null;
            }
        }

        public void Publica_en_SharePoint()
        {
            //office365.MS_Access msAccess = new office365.MS_Access();

            //EndesaBusiness.utilidades.ZIP zip = new utilidades.ZIP();
            utilidades.ZipUnZip zip7z = new utilidades.ZipUnZip();

            try
            {
                

                if (ss_pp.GetFecha_FinProceso("Contratación", "Solicitudes ATR", "Publica archivo en SharePoint").Date
                    < DateTime.Now.Date)
                {

                    ss_pp.Update_Fecha_Inicio("Contratación", "Solicitudes ATR", "Publica archivo en SharePoint");

                    Console.WriteLine("Publica SolATRMT en SharePoint");
                    Console.WriteLine("==============================");

                    FileInfo archivo = new FileInfo(param.GetValue("dirJob", DateTime.Now, DateTime.Now)
                        + param.GetValue("prefixIn", DateTime.Now, DateTime.Now)
                        + ultimoDiaHabil.ToString("yyMMdd") + ".txt");

                    Console.WriteLine("Llamando a programa extractor");
                    DescargaArchivoSolAtrMT();

                    if (archivo.Exists)
                    {
                        if (archivo.Length > 0)
                        {
                            FileInfo archivo_zip_previo = new FileInfo(param.GetValue("dirJob", DateTime.Now, DateTime.Now)
                            + param.GetValue("fileOut", DateTime.Now, DateTime.Now));

                            if (archivo_zip_previo.Exists)
                                archivo_zip_previo.Delete();

                            archivo.MoveTo(param.GetValue("dirJob", DateTime.Now, DateTime.Now)
                            + param.GetValue("fileOut", DateTime.Now, DateTime.Now));

                            string archivo_zip = param.GetValue("dirJob", DateTime.Now, DateTime.Now)
                                + param.GetValue("prefixIn", DateTime.Now, DateTime.Now)
                                + ultimoDiaHabil.ToString("yyMMdd") + ".zip";

                            //zip.Comprmir(archivo.FullName, archivo_zip);
                            zip7z.ComprimirArchivo(archivo.FullName, archivo_zip);
                            archivo_zip_previo = new FileInfo(archivo_zip);

                            FileInfo archivo_destino = new FileInfo(param.GetValue("ruta_sharepoint") + archivo_zip_previo.Name);

                            if (archivo_destino.Exists)
                                archivo_destino.Delete();

                            archivo_zip_previo.CopyTo(archivo_destino.FullName);

                            NotificaEntregaSOLATRMT(archivo, "SharePoint");

                            //Borramos el archivo ZIP
                            archivo_zip_previo.Delete();

                            ss_pp.Update_Fecha_Fin("Contratación", "Solicitudes ATR", "Publica archivo en SharePoint");

                        }
                    }
                    else
                    {
                        ss_pp.Update_Comentario("Contratación", "Solicitudes ATR", "Publica archivo en SharePoint",
                            "No existe el archivo: " + archivo.Name);
                    }
                        
                }

                
                
            }
            catch(Exception ex)
            {
                ficheroLog.AddError("Publica_en_SharePoint: " + ex.Message);
            }
                      


        }

        public void Publica_en_FTP_SolATRMT()
        {
            EndesaBusiness.utilidades.ZIP zip = new utilidades.ZIP();
            utilidades.ZipUnZip zip7z = new utilidades.ZipUnZip();

            try
            {
                FileInfo archivo = new FileInfo(param.GetValue("dirJob", DateTime.Now, DateTime.Now)
                + param.GetValue("prefixIn", DateTime.Now, DateTime.Now)
                + ultimoDiaHabil.ToString("yyMMdd") + ".txt");

                DescargaArchivoSolAtrMT();

                if (archivo.Exists)
                {
                    if (archivo.Length > 0)
                    {
                        FileInfo archivo_zip_previo = new FileInfo(param.GetValue("dirJob", DateTime.Now, DateTime.Now)
                        + param.GetValue("fileOut", DateTime.Now, DateTime.Now));

                        if (archivo_zip_previo.Exists)
                            archivo_zip_previo.Delete();

                        archivo.MoveTo(param.GetValue("dirJob", DateTime.Now, DateTime.Now)
                        + param.GetValue("fileOut", DateTime.Now, DateTime.Now));

                        string archivo_zip = param.GetValue("dirJob", DateTime.Now, DateTime.Now)
                            + param.GetValue("prefixIn", DateTime.Now, DateTime.Now)
                            + ultimoDiaHabil.ToString("yyMMdd") + ".zip";

                        //zip.Comprmir(archivo.FullName, archivo_zip);
                        zip7z.ComprimirArchivo(archivo.FullName, archivo_zip);
                        archivo_zip_previo = new FileInfo(archivo_zip);

                        // Subimos al FTP el archivo zip
                        PublicaEnFTP(archivo_zip_previo);


                        //Borramos el archivo ZIP
                        archivo_zip_previo.Delete();
                    }
                }
                    
            }
            catch(Exception ex)
            {
                ficheroLog.AddError("Publica_en_FTP_SolATRMT: " + ex.Message);
            }

            
        }

        public void ProcesoCarga()
        {

            FileInfo archivo_zip_previo = new FileInfo(param.GetValue("dirJob", DateTime.Now, DateTime.Now)
                + param.GetValue("fileOut", DateTime.Now, DateTime.Now));

            try
            {
                                

                if (ss_pp.GetFecha_FinProceso("Contratación", "Solicitudes ATR", "Carga Extracción en BBDD").Date
                    < DateTime.Now.Date && (ss_pp.GetFecha_FinProceso("Contratación", "Solicitudes ATR", "Publica archivo en SharePoint")
                    > ss_pp.GetFecha_InicioProceso("Contratación", "Solicitudes ATR", "Publica archivo en SharePoint")))
                {

                    ss_pp.Update_Fecha_Inicio("Contratación", "Solicitudes ATR", "Carga Extracción en BBDD");

                    bool hayError = CargaSolATR_porLinea_PosFijas(archivo_zip_previo.FullName);
                    if (!hayError)
                    {
                        hayError = Vuelca_a_MySQL(lista);

                        if (!hayError)
                        {
                            CargaSolAtr(null);
                            ss_pp.Update_Fecha_Fin("Contratación", "Solicitudes ATR", "Carga Extracción en BBDD");

                            // Llamamos a la cola de Contratación
                            
                            ss_pp.Update_Fecha_Inicio("Contratación", "Cola Procesos Contratación", "Cola Procesos Contratación");
                            
                            utilidades.Fichero.EjecutaComando(param.GetValue("ruta_cola_procesos"), 
                                param.GetValue("nombre_cola_procesos"));

                            ss_pp.Update_Fecha_Fin("Contratación", "Cola Procesos Contratación", "Cola Procesos Contratación");
                        }
                        else
                        {

                        }

                    }
                    else
                    {

                    }

                    
                }
            }catch(Exception e)
            {
                ficheroLog.AddError("ProcesoCarga: " + e.Message);
            }


           
        }


        private bool Vuelca_a_MySQL(List<EndesaEntity.contratacion.SolATRMT> lista)
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int i = 0;
            int total_registros = 0;
            string strSql = "";

            try
            {
                Console.WriteLine("");

                strSql = "delete from solatrmt_tmp";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                foreach (EndesaEntity.contratacion.SolATRMT c in lista)
                {
                    i++;
                    total_registros++;
                    if (firstOnly)
                    {
                        sb.Append("replace into solatrmt_tmp");
                        sb.Append(" (EMP_TIT, estadoSolATR, codSolATR, tipoSolATR, fRecepcion, fAcepRech, fRechazo,");
                        sb.Append("fEnvioATR, fEnvioDoc, fPrevAlta, fActivacion, CCOUNIPS, CUPS_EXT, TIPO_CLI,");
                        sb.Append("CIF, cliente, SEG_MER, LINEA_N, DIREC_PS, CDISTRIB, CONTR_ATR,");
                        sb.Append("COMENT_1_SOLICITUD, COMENT_2_SOLICITUD, COMENT_3_SOLICITUD, CHECK_ANULAC, MOT_BAJA, POTENCIA1, POTENCIA2,");
                        sb.Append("POTENCIA3, POTENCIA4, POTENCIA5, POTENCIA6, TARIFA, TENSION, CCONTRPS,");
                        sb.Append("VER_CONTR_PS, USO_CONTR, GESTOR_SCE, COMENT_1_PETTRAB, COMENT_2_PETTRAB, COMENT_3_PETTRAB, MOT_RECH_1,");
                        sb.Append("COMENT_1_MOTRECH_1, COMENT_2_MOTRECH_1, COMENT_3_MOTRECH_1, MOT_RECH_2, COMENT_1_MOTRECH_2, COMENT_2_MOTRECH_2, COMENT_3_MOTRECH_2,");
                        sb.Append("MOT_RECH_3, COMENT_1_MOTRECH_3, COMENT_2_MOTRECH_3, COMENT_3_MOTRECH_3, MOT_RECH_4, COMENT_1_MOTRECH_4, COMENT_2_MOTRECH_4,");
                        sb.Append("COMENT_3_MOTRECH_4, MOT_RECH_5, COMENT_1_MOTRECH_5, COMENT_2_MOTRECH_5, COMENT_3_MOTRECH_5, MOT_RECH_6, COMENT_1_MOTRECH_6,");
                        sb.Append("COMENT_2_MOTRECH_6, COMENT_3_MOTRECH_6, MOT_RECH_7, COMENT_1_MOTRECH_7, COMENT_2_MOTRECH_7, COMENT_3_MOTRECH_7, MOT_RECH_8,");
                        sb.Append("COMENT_1_MOTRECH_8, COMENT_2_MOTRECH_8, COMENT_3_MOTRECH_8, MOT_RECH_9, COMENT_1_MOTRECH_9, COMENT_2_MOTRECH_9, COMENT_3_MOTRECH_9,");
                        sb.Append("MOT_RECH_10, COMENT_1_MOTRECH_10, COMENT_2_MOTRECH_10, COMENT_3_MOTRECH_10, TESTCONT, MANUAL, MED_BAJ,");
                        sb.Append("PT_CONT, PRC_PERD, MOD_FECHA_SOLCT, FECHA_SOLCT, MOD_FECHA_RESP, FECHA_RESP) values ");
                        firstOnly = false;
                    }

                    sb.Append("(").Append(c.empresa_titular).Append(",");
                    sb.Append("'").Append(c.estadoSolAtr).Append("',");
                    sb.Append(c.codSolAtr).Append(",");
                    sb.Append("'").Append(c.tipoSolAtr).Append("',");

                    sb.Append(c.fRecepcion).Append(",");
                    sb.Append(c.fAcepRech).Append(",");
                    sb.Append(c.fRechazo).Append(",");
                    sb.Append(c.fEnvioAtr).Append(",");
                    sb.Append(c.fEnvioDoc).Append(",");
                    sb.Append(c.fPrevAlta).Append(",");
                    sb.Append(c.fActivacion).Append(",");

                    sb.Append("'").Append(c.ccounips).Append("',");
                    sb.Append("'").Append(c.cups_ext).Append("',");
                    sb.Append("'").Append(c.tipo_cli).Append("',");
                    sb.Append("'").Append(c.cif).Append("',");
                    sb.Append("'").Append(c.cliente).Append("',");
                    sb.Append("'").Append(c.seg_mer).Append("',");
                    sb.Append(c.linea_n).Append(",");
                    sb.Append("'").Append(c.direc_ps).Append("',");
                    sb.Append("'").Append(c.cdistrib).Append("',");
                    sb.Append(c.contr_atr).Append(",");
                    
                    sb.Append("'").Append(c.coment_1_solicitud).Append("',");
                    sb.Append("'").Append(c.coment_2_solicitud).Append("',");
                    sb.Append("'").Append(c.coment_3_solicitud).Append("',");

                    sb.Append("'").Append(c.check_anulac).Append("',");
                    sb.Append("'").Append(c.mot_baja).Append("',");

                    for(int j = 0; j < 6; j++)                    
                        sb.Append(c.potencia[j]).Append(",");


                    sb.Append("'").Append(c.tarifa).Append("',");
                    sb.Append(c.tension).Append(",");
                    sb.Append("'").Append(c.ccontrps).Append("',");
                    sb.Append(c.ver_contr_ps).Append(",");
                    sb.Append("'").Append(c.uso_contr).Append("',");
                    sb.Append("'").Append(c.gestor_sce).Append("',");

                    sb.Append("'").Append(c.coment_1_pettrab).Append("',");
                    sb.Append("'").Append(c.coment_2_pettrab).Append("',");
                    sb.Append("'").Append(c.coment_3_pettrab).Append("',");


                    for (int j = 0; j < 10; j++)
                    {
                        
                        sb.Append("'").Append(c.mot_rech[j]).Append("',");
                        sb.Append("'").Append(c.coment_1_motrech[j]).Append("',");
                        sb.Append("'").Append(c.coment_2_motrech[j]).Append("',");
                        sb.Append("'").Append(c.coment_3_motrech[j]).Append("',");

                    }

                    sb.Append(c.testcont).Append(",");
                    sb.Append("'").Append(c.manual).Append("',");
                    sb.Append("'").Append(c.med_baj).Append("',");
                    sb.Append(c.pt_cont).Append(",");
                    sb.Append(c.prc_perd).Append(",");
                    sb.Append("'").Append(c.mod_fecha_solct).Append("',");
                    sb.Append(c.fecha_solct).Append(",");
                    sb.Append("'").Append(c.mod_fecha_resp).Append("',");
                    sb.Append(c.fecha_resp).Append("),");



                    if (i == 250)
                    {
                        Console.CursorLeft = 0;
                        Console.Write(total_registros);
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        i = 0;
                    }
                }

                if (i > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                }

                return false;

            }
            catch(Exception e)
            {
                ficheroLog.AddError("Vuelca_a_MySQL: " + e.Message);
                return true;
            }
        }

        private void DescargaArchivoSolAtrMT()
        {
            ficheroLog.Add("Ejecutando extractor: " + param.GetValue("extractor", DateTime.Now, DateTime.Now));
            utilidades.Fichero.EjecutaComando(param.GetValue("extractor", DateTime.Now, DateTime.Now), ultimoDiaHabil.ToString("yyMMdd"));

        }

        public bool CargaSolATR_porLinea(string fichero)
        {

            string[] campos;
            int total_registros = 0;
            int c = 0;           
            string line = "";
            int i = 0;

            

            try
            {

                System.IO.StreamReader file = new System.IO.StreamReader(fichero, System.Text.Encoding.GetEncoding(1252));
                while ((line = file.ReadLine()) != null)
                {
                    i++;

                    if(i > 1)
                    {
                        total_registros++;
                        campos = line.Split(';');
                        if(campos.Length == 91)
                        {
                            Console.CursorLeft = 0;
                            Console.Write("Importando línea: " + i);

                            c = 0;

                            EndesaEntity.contratacion.SolATRMT cc = new EndesaEntity.contratacion.SolATRMT();

                            cc.empresa_titular = Convert.ToInt32(campos[c]); c++;
                            cc.estadoSolAtr = utilidades.FuncionesTexto.RT(campos[c]); c++;
                            cc.codSolAtr = Convert.ToInt64(campos[c]); c++;
                            cc.tipoSolAtr = utilidades.FuncionesTexto.RT(campos[c]); c++;

                            if (campos[c].Trim() != "")
                                cc.fRecepcion = Convert.ToInt32(campos[c]); c++;

                            if (campos[c].Trim() != "")
                                cc.fAcepRech = Convert.ToInt32(campos[c]); c++;

                            if (campos[c].Trim() != "")
                                cc.fRechazo = Convert.ToInt32(campos[c]); c++;

                            if (campos[c].Trim() != "")
                                cc.fEnvioAtr = Convert.ToInt32(campos[c]); c++;

                            if (campos[c].Trim() != "")
                                cc.fEnvioDoc = Convert.ToInt32(campos[c]); c++;

                            if (campos[c].Trim() != "")
                                cc.fPrevAlta = Convert.ToInt32(campos[c]); c++;

                            if (campos[c].Trim() != "")
                                cc.fActivacion = Convert.ToInt32(campos[c]); c++;

                            cc.ccounips = utilidades.FuncionesTexto.RT(campos[c]); c++;
                            cc.cups_ext = utilidades.FuncionesTexto.RT(campos[c]); c++;
                            cc.tipo_cli = utilidades.FuncionesTexto.RT(campos[c]); c++;
                            cc.cif = utilidades.FuncionesTexto.RT(campos[c]); c++;
                            cc.cliente = utilidades.FuncionesTexto.RT(campos[c]); c++;
                            cc.seg_mer = utilidades.FuncionesTexto.RT(campos[c]); c++;

                            if (campos[c].Trim() != "")
                                cc.linea_n = Convert.ToInt32(campos[c]); c++;

                            cc.direc_ps = utilidades.FuncionesTexto.RT(campos[c]); c++;
                            cc.cdistrib = utilidades.FuncionesTexto.RT(campos[c]); c++;

                            if (campos[c].Trim() != "")
                                cc.contr_atr = Convert.ToInt64(campos[c]); c++;

                            cc.coment_1_solicitud = utilidades.FuncionesTexto.RT(campos[c].Replace("\\\"", "").Replace("+", "")); c++;
                            cc.coment_2_solicitud = utilidades.FuncionesTexto.RT(campos[c].Replace("\\\"", "").Replace("+", "")); c++;
                            cc.coment_3_solicitud = utilidades.FuncionesTexto.RT(campos[c].Replace("\\\"", "").Replace("+", "")); c++;
                            cc.check_anulac = campos[c].Trim(); c++;
                            cc.mot_baja = utilidades.FuncionesTexto.RT(campos[c]); c++;

                            for (int j = 0; j <6; j++)
                            {
                                if(campos[c].Trim() != "" && campos[c].Trim() != "00000000.000")
                                    cc.potencia[j] = Convert.ToDouble(campos[c]); 
                                c++;
                            }

                            cc.tarifa = campos[c].Trim(); c++;
                            cc.tension = Convert.ToInt32(campos[c]); c++;
                            cc.ccontrps = campos[c].Trim(); c++;
                            cc.ver_contr_ps = Convert.ToInt32(campos[c]); c++;
                            cc.uso_contr = campos[c].Trim(); c++;
                            cc.gestor_sce = utilidades.FuncionesTexto.RT(campos[c]); c++;
                            cc.coment_1_pettrab = utilidades.FuncionesTexto.RT(campos[c].Replace("\\\"", "").Replace("+","")); c++;
                            cc.coment_2_pettrab = utilidades.FuncionesTexto.RT(campos[c].Replace("\\\"", "").Replace("+", "")); c++;
                            cc.coment_3_pettrab = utilidades.FuncionesTexto.RT(campos[c].Replace("\\\"", "").Replace("+", "")); c++;

                            for (int j = 0; j < 10; j++)
                            {
                                cc.mot_rech[j] = utilidades.FuncionesTexto.RT(campos[c].Replace("\\\"", "").Replace("+", "")); c++;
                                cc.coment_1_motrech[j] = utilidades.FuncionesTexto.RT(campos[c].Replace("\\\"", "").Replace("+", "")); c++;
                                cc.coment_2_motrech[j] = utilidades.FuncionesTexto.RT(campos[c].Replace("\\\"", "").Replace("+", "")); c++;
                                cc.coment_3_motrech[j] = utilidades.FuncionesTexto.RT(campos[c].Replace("\\\"", "").Replace("+", "")); c++;
                            }

                            if (campos[c].Trim() != "")
                                cc.testcont = Convert.ToInt32(campos[c]); c++;

                            cc.manual = campos[c].Trim(); c++;
                            cc.med_baj = campos[c].Trim(); c++;

                            if (campos[c].Trim() != "")
                                cc.pt_cont = Convert.ToInt64(campos[c]); c++;

                            if (campos[c].Trim() != "" && campos[c].Trim() != "000.000000")
                                cc.prc_perd = Convert.ToDouble(campos[c]); c++;

                            cc.mod_fecha_solct = utilidades.FuncionesTexto.RT(campos[c]); c++;

                            if (campos[c].Trim() != "")
                                cc.fecha_solct = Convert.ToInt32(campos[c]); c++;

                            cc.mod_fecha_resp = utilidades.FuncionesTexto.RT(campos[c]); c++;

                            if (campos[c].Trim() != "")
                                cc.fecha_resp = Convert.ToInt32(campos[c]); c++;

                            lista.Add(cc);

                        }
                        else
                        {
                            // Error en fichero
                            Console.WriteLine("Error en la línea: " + i + " y columna: " + c);
                            Console.WriteLine(line);
                            ficheroLog.AddError(line);
                                
                        }
                       

                    }

                   

                }
                file.Close();

                return false;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return true;
            }
        }

        public bool CargaSolATR_porLinea_PosFijas(string fichero)
        {
                        
            int total_registros = 0;
            int c = 0;
            int p1 = 0;
            int p2 = 0;
            string line = "";
            int i = 0;

            try
            {
                

                System.IO.StreamReader file = new System.IO.StreamReader(fichero, System.Text.Encoding.GetEncoding(1252));
                while ((line = file.ReadLine()) != null)
                {
                    i++;

                    if (i > 1)
                    {
                        total_registros++;                       
                       
                        Console.CursorLeft = 0;
                        Console.Write("Importando línea: " + i);

                        c = 0;

                        EndesaEntity.contratacion.SolATRMT cc = new EndesaEntity.contratacion.SolATRMT();

                        cc.empresa_titular = Convert.ToInt32(Cad(line,1,6)); 
                        cc.estadoSolAtr = utilidades.FuncionesTexto.RT(Cad(line, 7, 37));
                        cc.codSolAtr = Convert.ToInt64(Cad(line, 38, 50)); 
                        cc.tipoSolAtr = utilidades.FuncionesTexto.RT(Cad(line, 51, 81));                       


                        if (Cad(line, 82, 90) != "")
                            cc.fRecepcion = Convert.ToInt32(Cad(line, 82, 90));

                        if (Cad(line, 91, 99) != "")
                            cc.fAcepRech = Convert.ToInt32(Cad(line, 91, 99)); 

                        if (Cad(line, 100, 108) != "")
                            cc.fRechazo = Convert.ToInt32(Cad(line, 100, 108));

                        if (Cad(line, 109, 117) != "")
                            cc.fEnvioAtr = Convert.ToInt32(Cad(line, 109, 117)); 

                        if (Cad(line, 118, 126) != "")
                            cc.fEnvioDoc = Convert.ToInt32(Cad(line, 118, 126)); 

                        if (Cad(line, 127, 135) != "")
                            cc.fPrevAlta = Convert.ToInt32(Cad(line, 127, 135)); 

                        if (Cad(line, 136, 144) != "")
                            cc.fActivacion = Convert.ToInt32(Cad(line, 136, 144)); 

                        cc.ccounips = utilidades.FuncionesTexto.RT(Cad(line, 145, 158)); 
                        cc.cups_ext = utilidades.FuncionesTexto.RT(Cad(line, 159, 181)); 
                        cc.tipo_cli = utilidades.FuncionesTexto.RT(Cad(line, 182, 183)); 
                        cc.cif = utilidades.FuncionesTexto.RT(Cad(line, 184, 199)); 
                        cc.cliente = utilidades.FuncionesTexto.RT(Cad(line, 200, 240)); 
                        cc.seg_mer = utilidades.FuncionesTexto.RT(Cad(line, 241, 242)); 

                        if (Cad(line, 243, 246) != "")
                            cc.linea_n = Convert.ToInt32(Cad(line, 243, 246)); 

                        cc.direc_ps = utilidades.FuncionesTexto.RT(Cad(line, 247, 327)); 
                        cc.cdistrib = utilidades.FuncionesTexto.RT(Cad(line, 328, 331)); 

                        if (Cad(line, 332, 344) != "")
                            cc.contr_atr = Convert.ToInt64(Cad(line, 332, 344)); 

                        cc.coment_1_solicitud = utilidades.FuncionesTexto.RT(Cad(line, 345, 467).Replace("\\\"", "").Replace("+", "")); 
                        cc.coment_2_solicitud = utilidades.FuncionesTexto.RT(Cad(line, 468, 590).Replace("\\\"", "").Replace("+", "")); 
                        cc.coment_3_solicitud = utilidades.FuncionesTexto.RT(Cad(line, 591, 713).Replace("\\\"", "").Replace("+", "")); 
                        cc.check_anulac = Cad(line, 714, 716); 
                        cc.mot_baja = utilidades.FuncionesTexto.RT(Cad(line, 717, 719)); 

                           
                        if (Cad(line, 720, 732) != "" && Cad(line, 720, 732) != "00000000.000")
                            cc.potencia[0] = Convert.ToDouble(Cad(line, 720, 732));

                        if (Cad(line, 733, 745) != "" && Cad(line, 733, 745) != "00000000.000")
                            cc.potencia[1] = Convert.ToDouble(Cad(line, 733, 745));

                        if (Cad(line, 746, 758) != "" && Cad(line, 746, 758) != "00000000.000")
                            cc.potencia[2] = Convert.ToDouble(Cad(line, 746, 758));

                        if (Cad(line, 759, 771) != "" && Cad(line, 759, 771) != "00000000.000")
                            cc.potencia[3] = Convert.ToDouble(Cad(line, 759, 771));

                        if (Cad(line, 772, 784) != "" && Cad(line, 772, 784) != "00000000.000")
                            cc.potencia[4] = Convert.ToDouble(Cad(line, 772, 784));

                        if (Cad(line, 785, 797) != "" && Cad(line, 785, 797) != "00000000.000")
                            cc.potencia[5] = Convert.ToDouble(Cad(line, 785, 797));



                        cc.tarifa = Cad(line, 798, 809); 
                        cc.tension = Convert.ToInt32(Cad(line, 810, 818)); 
                        cc.ccontrps = Cad(line, 819, 831); 
                        cc.ver_contr_ps = Convert.ToInt32(Cad(line, 832, 835)); 
                        cc.uso_contr = Cad(line, 836, 837); 
                        cc.gestor_sce = utilidades.FuncionesTexto.RT(Cad(line, 838, 878)); 
                        cc.coment_1_pettrab = utilidades.FuncionesTexto.RT(Cad(line, 879, 1001).Replace("\\\"", "").Replace("+", "")); 
                        cc.coment_2_pettrab = utilidades.FuncionesTexto.RT(Cad(line, 1002, 1124).Replace("\\\"", "").Replace("+", "")); 
                        cc.coment_3_pettrab = utilidades.FuncionesTexto.RT(Cad(line, 1125, 1247).Replace("\\\"", "").Replace("+", ""));


                        p1 = 1247; 
                        for (int j = 0; j < 10; j++)
                        {
                            p1++;
                            p2 = p1 + 30;
                            cc.mot_rech[j] = utilidades.FuncionesTexto.RT(Cad(line, p1, p2).Replace("\\\"", "").Replace("+", ""));
                            p1 = p2 + 1;
                            p2 = p1 + 122;
                            cc.coment_1_motrech[j] = utilidades.FuncionesTexto.RT(Cad(line, p1, p2).Replace("\\\"", "").Replace("+", ""));
                            p1 = p2 + 1;
                            p2 = p1 + 122;
                            cc.coment_2_motrech[j] = utilidades.FuncionesTexto.RT(Cad(line, p1, p2).Replace("\\\"", "").Replace("+", ""));
                            p1 = p2 + 1;
                            p2 = p1 + 122;
                            cc.coment_3_motrech[j] = utilidades.FuncionesTexto.RT(Cad(line, p1, p2).Replace("\\\"", "").Replace("+", ""));
                        }

                        if (Cad(line, 5248, 5251) != "")
                            cc.testcont = Convert.ToInt32(Cad(line, 5248, 5251)); 

                        cc.manual = Cad(line, 5252, 5253); 
                        cc.med_baj = Cad(line, 5254, 5255);

                        if (Cad(line, 5256, 5270) != "")
                            cc.pt_cont = Convert.ToInt64(Cad(line, 5256, 5270)); 

                        if (Cad(line, 5271, 5281) != "" && Cad(line, 5271, 5281) != "000.000000")
                            cc.prc_perd = Convert.ToDouble(Cad(line, 5271, 5281)); 

                        cc.mod_fecha_solct = utilidades.FuncionesTexto.RT(Cad(line, 5282, 5284)); 

                        if (Cad(line, 5285, 5293) != "")
                            cc.fecha_solct = Convert.ToInt32(Cad(line, 5285, 5293)); 

                        cc.mod_fecha_resp = utilidades.FuncionesTexto.RT(Cad(line, 5294, 5296)); 

                        if (Cad(line, 5297, 5305) != "")
                            cc.fecha_resp = Convert.ToInt32(Cad(line, 5297, 5305)); 

                        lista.Add(cc);

                    }

                }
                file.Close();



                return false;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("CargaSolATR_porLinea_PosFijas: " + e.Message + "  --> línea: " + i);
                return true;
            }
        }

        private string Cad(string linea, int posini, int posfin)
        {
            string cadena;


            cadena = linea.Substring(posini - 1, posfin - posini);
            cadena = cadena.Trim();
            return cadena;
        }

        public void CargaSolAtr(string fichero)
        {

            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {

                strSql = "delete from SOLATRMT_ANT";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "replace into SOLATRMT_ANT select * from SOLATRMT";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "delete from SOLATRMT";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "replace into SOLATRMT select * from solatrmt_tmp";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                #region LOAD DATA LOCAL INFILE
                //strSql = "LOAD DATA LOCAL INFILE '" + fichero.Replace(@"\", "\\\\")
                //    + "' REPLACE INTO TABLE SOLATRMT "
                //    + "FIELDS TERMINATED BY ';' LINES TERMINATED BY '\\n' "
                //    + "IGNORE 1 LINES "
                //    + "(@EMP_TIT,@estadoSolATR,@codSolATR,@tipoSolATR,@fRecepcion,@fAcepRech,@fRechazo,@fEnvioATR,@fEnvioDoc,"
                //    + "@fPrevAlta,@fActivacion,@CCOUNIPS,@CUPS_EXT,@TIPO_CLI,@CIF,@cliente,@SEG_MER,@LINEA_N,@DIREC_PS,"
                //    + "@CDISTRIB,@CONTR_ATR,@COMENT_1_SOLICITUD,@COMENT_2_SOLICITUD,@COMENT_3_SOLICITUD,@CHECK_ANULAC,"
                //    + "@MOT_BAJA,@POTENCIA1,@POTENCIA2,@POTENCIA3,@POTENCIA4,@POTENCIA5,@POTENCIA6,@TARIFA,@TENSION,"
                //    + "@CCONTRPS,@VER_CONTR_PS,@USO_CONTR,@GESTOR_SCE,@COMENT_1_PETTRAB,@COMENT_2_PETTRAB,@COMENT_3_PETTRAB,"
                //    + "@MOT_RECH_1,@COMENT_1_MOTRECH_1,@COMENT_2_MOTRECH_1,@COMENT_3_MOTRECH_1,@MOT_RECH_2,@COMENT_1_MOTRECH_2,"
                //    + "@COMENT_2_MOTRECH_2,@COMENT_3_MOTRECH_2,@MOT_RECH_3,@COMENT_1_MOTRECH_3,@COMENT_2_MOTRECH_3,"
                //    + "@COMENT_3_MOTRECH_3,@MOT_RECH_4,@COMENT_1_MOTRECH_4,@COMENT_2_MOTRECH_4,@COMENT_3_MOTRECH_4,"
                //    + "@MOT_RECH_5,@COMENT_1_MOTRECH_5,@COMENT_2_MOTRECH_5,@COMENT_3_MOTRECH_5,@MOT_RECH_6,@COMENT_1_MOTRECH_6,"
                //    + "@COMENT_2_MOTRECH_6,@COMENT_3_MOTRECH_6,@MOT_RECH_7,@COMENT_1_MOTRECH_7,@COMENT_2_MOTRECH_7,"
                //    + "@COMENT_3_MOTRECH_7,@MOT_RECH_8,@COMENT_1_MOTRECH_8,@COMENT_2_MOTRECH_8,@COMENT_3_MOTRECH_8,"
                //    + "@MOT_RECH_9,@COMENT_1_MOTRECH_9,@COMENT_2_MOTRECH_9,@COMENT_3_MOTRECH_9,@MOT_RECH_10,"
                //    + "@COMENT_1_MOTRECH_10,@COMENT_2_MOTRECH_10,@COMENT_3_MOTRECH_10,@TESTCONT,@MANUAL,@MED_BAJ,"
                //    + "@PT_CONT,@PRC_PERD,@MOD_FECHA_SOLCT,@FECHA_SOLCT,@MOD_FECHA_RESP,@FECHA_RESP) SET "
                //    + "EMP_TIT = @EMP_TIT, "
                //    + "estadoSolATR = trim(@estadoSolATR), "
                //    + "codSolATR = @codSolATR, "
                //    + "tipoSolATR =  trim(replace(@tipoSolATR,'" + @"""" + "', '')), "
                //    + "fRecepcion = @fRecepcion, "
                //    + "fAcepRech = @fAcepRech, "
                //    + "fRechazo = @fRechazo, "
                //    + "fEnvioATR = @fEnvioATR, "
                //    + "fEnvioDoc = @fEnvioDoc, "
                //    + "fPrevAlta = @fPrevAlta, "
                //    + "fActivacion = @fActivacion, "
                //    + "CCOUNIPS = trim(@CCOUNIPS), "
                //    + "CUPS_EXT = trim(@CUPS_EXT), "
                //    + "TIPO_CLI = @TIPO_CLI, "
                //    + "CIF = trim(@CIF), "
                //    + "cliente = trim(@cliente), "
                //    + "SEG_MER = @SEG_MER, "
                //    + "LINEA_N = @LINEA_N, "
                //    + "DIREC_PS = trim(@DIREC_PS), "
                //    + "CDISTRIB = trim(@CDISTRIB), "
                //    + "CONTR_ATR = @CONTR_ATR, "
                //    + "COMENT_1_SOLICITUD = trim(replace(@COMENT_1_SOLICITUD,'" + @"""" + "', '')), "
                //    + "COMENT_2_SOLICITUD = trim(replace(@COMENT_2_SOLICITUD,'" + @"""" + "', '')), "
                //    + "COMENT_3_SOLICITUD = trim(replace(@COMENT_3_SOLICITUD,'" + @"""" + "', '')), "
                //    + "CHECK_ANULAC = @CHECK_ANULAC, "
                //    + "MOT_BAJA = @MOT_BAJA, "
                //    + "POTENCIA1 = @POTENCIA1, "
                //    + "POTENCIA2 = @POTENCIA2, "
                //    + "POTENCIA3 = @POTENCIA3, "
                //    + "POTENCIA4 = @POTENCIA4, "
                //    + "POTENCIA5 = @POTENCIA5, "
                //    + "POTENCIA6 = @POTENCIA6, "
                //    + "TARIFA = @TARIFA, "
                //    + "TENSION = @TENSION, "
                //    + "CCONTRPS = @CCONTRPS, "
                //    + "VER_CONTR_PS = @VER_CONTR_PS, "
                //    + "USO_CONTR = @USO_CONTR, "
                //    + "GESTOR_SCE = @GESTOR_SCE, "
                //    + "COMENT_1_PETTRAB = trim(replace(@COMENT_1_PETTRAB,'" + @"""" + "', '')), "
                //    + "COMENT_2_PETTRAB = trim(replace(@COMENT_2_PETTRAB,'" + @"""" + "', '')), "
                //    + "COMENT_3_PETTRAB = trim(replace(@COMENT_3_PETTRAB,'" + @"""" + "', '')), "
                //    + "MOT_RECH_1 = trim(replace(@MOT_RECH_1,'" + @"""" + "', '')), "
                //    + "COMENT_1_MOTRECH_1 = trim(replace(@COMENT_1_MOTRECH_1,'" + @"""" + "', '')), "
                //    + "COMENT_2_MOTRECH_1 = trim(replace(@COMENT_2_MOTRECH_1,'" + @"""" + "', '')), "
                //    + "COMENT_3_MOTRECH_1 = trim(replace(@COMENT_3_MOTRECH_1,'" + @"""" + "', '')), "
                //    + "MOT_RECH_2 = trim(replace(@MOT_RECH_2,'" + @"""" + "', '')), "
                //    + "COMENT_1_MOTRECH_2 = trim(replace(@COMENT_1_MOTRECH_2,'" + @"""" + "', '')), "
                //    + "COMENT_2_MOTRECH_2 = trim(replace(@COMENT_2_MOTRECH_2,'" + @"""" + "', '')), "
                //    + "COMENT_3_MOTRECH_2 = trim(replace(@COMENT_3_MOTRECH_2,'" + @"""" + "', '')), "
                //    + "MOT_RECH_3 = trim(replace(@MOT_RECH_3,'" + @"""" + "', '')), "
                //    + "COMENT_1_MOTRECH_3 = trim(replace(@COMENT_1_MOTRECH_3,'" + @"""" + "', '')), "
                //    + "COMENT_2_MOTRECH_3 = trim(replace(@COMENT_2_MOTRECH_3,'" + @"""" + "', '')), "
                //    + "COMENT_3_MOTRECH_3 = trim(replace(@COMENT_3_MOTRECH_3,'" + @"""" + "', '')), "
                //    + "MOT_RECH_4 = trim(replace(@MOT_RECH_4,'" + @"""" + "', '')), "
                //    + "COMENT_1_MOTRECH_4 = trim(replace(@COMENT_1_MOTRECH_4,'" + @"""" + "', '')), "
                //    + "COMENT_2_MOTRECH_4 = trim(replace(@COMENT_2_MOTRECH_4,'" + @"""" + "', '')), "
                //    + "COMENT_3_MOTRECH_4 = trim(replace(@COMENT_3_MOTRECH_4,'" + @"""" + "', '')), "
                //    + "MOT_RECH_5 = trim(replace(@MOT_RECH_5,'" + @"""" + "', '')), "
                //    + "COMENT_1_MOTRECH_5 = trim(replace(@COMENT_1_MOTRECH_5,'" + @"""" + "', '')), "
                //    + "COMENT_2_MOTRECH_5 = trim(replace(@COMENT_2_MOTRECH_5,'" + @"""" + "', '')), "
                //    + "COMENT_3_MOTRECH_5 = trim(replace(@COMENT_3_MOTRECH_5,'" + @"""" + "', '')), "
                //    + "MOT_RECH_6 = trim(replace(@MOT_RECH_6,'" + @"""" + "', '')), "
                //    + "COMENT_1_MOTRECH_6 = trim(replace(@COMENT_1_MOTRECH_6,'" + @"""" + "', '')), "
                //    + "COMENT_2_MOTRECH_6 = trim(replace(@COMENT_2_MOTRECH_6,'" + @"""" + "', '')), "
                //    + "COMENT_3_MOTRECH_6 = trim(replace(@COMENT_3_MOTRECH_6,'" + @"""" + "', '')), "
                //    + "MOT_RECH_7 = trim(replace(@MOT_RECH_7,'" + @"""" + "', '')), "
                //    + "COMENT_1_MOTRECH_7 = trim(replace(@COMENT_1_MOTRECH_7,'" + @"""" + "', '')), "
                //    + "COMENT_2_MOTRECH_7 = trim(replace(@COMENT_2_MOTRECH_7,'" + @"""" + "', '')), "
                //    + "COMENT_3_MOTRECH_7 = trim(replace(@COMENT_3_MOTRECH_7,'" + @"""" + "', '')), "
                //    + "MOT_RECH_8 = trim(replace(@MOT_RECH_8,'" + @"""" + "', '')), "
                //    + "COMENT_1_MOTRECH_8 = trim(replace(@COMENT_1_MOTRECH_8,'" + @"""" + "', '')), "
                //    + "COMENT_2_MOTRECH_8 = trim(replace(@COMENT_2_MOTRECH_8,'" + @"""" + "', '')), "
                //    + "COMENT_3_MOTRECH_8 = trim(replace(@COMENT_3_MOTRECH_8,'" + @"""" + "', '')), "
                //    + "MOT_RECH_9 = trim(replace(@MOT_RECH_9,'" + @"""" + "', '')), "
                //    + "COMENT_1_MOTRECH_9 = trim(replace(@COMENT_1_MOTRECH_9,'" + @"""" + "', '')), "
                //    + "COMENT_2_MOTRECH_9 = trim(replace(@COMENT_2_MOTRECH_9,'" + @"""" + "', '')), "
                //    + "COMENT_3_MOTRECH_9 = trim(replace(@COMENT_3_MOTRECH_9,'" + @"""" + "', '')), "
                //    + "MOT_RECH_10 = trim(replace(@MOT_RECH_10,'" + @"""" + "', '')), "
                //    + "COMENT_1_MOTRECH_10 = trim(replace(@COMENT_1_MOTRECH_10,'" + @"""" + "', '')), "
                //    + "COMENT_2_MOTRECH_10 = trim(replace(@COMENT_2_MOTRECH_10,'" + @"""" + "', '')), "
                //    + "COMENT_3_MOTRECH_10 = trim(replace(@COMENT_3_MOTRECH_10,'" + @"""" + "', '')), "
                //    + "TESTCONT = @TESTCONT, "
                //    + "MANUAL = @MANUAL, "
                //    + "MED_BAJ = @MED_BAJ, "
                //    + "PT_CONT = @PT_CONT, "
                //    + "PRC_PERD = @PRC_PERD, "
                //    + "MOD_FECHA_SOLCT = @MOD_FECHA_SOLCT, "
                //    + "FECHA_SOLCT = @FECHA_SOLCT, "
                //    + "MOD_FECHA_RESP = @MOD_FECHA_RESP, "
                //    + "FECHA_RESP = @FECHA_RESP; ";

                //db = new MySQLDB(MySQLDB.Esquemas.CON);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                #endregion

                ficheroLog.Add("");
                ficheroLog.Add("Genera_SOLATRMT_MAX");
                ficheroLog.Add("===================");

                Genera_SOLATRMT_MAX();

                strSql = "DELETE FROM SOLATRMT_MAX_ANT";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "replace into SOLATRMT_MAX_ANT select a.EMP_TIT,a.estadoSolATR,max(a.codSolATR) codSolATR,"
                    + "a.tipoSolATR,a.fRecepcion,a.fAcepRech,"
                    + "a.fRechazo,a.fEnvioATR,a.fEnvioDoc,"
                    + "a.fPrevAlta,a.fActivacion,a.CCOUNIPS,"
                    + "a.CUPS_EXT,a.TIPO_CLI,a.CIF,"
                    + "a.cliente,a.SEG_MER,a.LINEA_N,"
                    + "a.DIREC_PS,a.CDISTRIB,a.CONTR_ATR,"
                    + "a.COMENT_1_SOLICITUD,a.COMENT_2_SOLICITUD,a.COMENT_3_SOLICITUD,"
                    + "a.CHECK_ANULAC,a.MOT_BAJA,a.POTENCIA1,"
                    + "a.POTENCIA2,a.POTENCIA3,a.POTENCIA4,"
                    + "a.POTENCIA5,a.POTENCIA6,a.TARIFA,"
                    + "a.TENSION,a.CCONTRPS,a.VER_CONTR_PS,"
                    + "a.USO_CONTR,a.GESTOR_SCE,a.COMENT_1_PETTRAB,"
                    + "a.COMENT_2_PETTRAB,a.COMENT_3_PETTRAB,a.MOT_RECH_1,"
                    + "a.COMENT_1_MOTRECH_1,a.COMENT_2_MOTRECH_1,a.COMENT_3_MOTRECH_1,"
                    + "a.MOT_RECH_2,a.COMENT_1_MOTRECH_2,a.COMENT_2_MOTRECH_2,"
                    + "a.COMENT_3_MOTRECH_2,a.MOT_RECH_3,a.COMENT_1_MOTRECH_3,"
                    + "a.COMENT_2_MOTRECH_3,a.COMENT_3_MOTRECH_3,a.MOT_RECH_4,"
                    + "a.COMENT_1_MOTRECH_4,a.COMENT_2_MOTRECH_4,a.COMENT_3_MOTRECH_4,"
                    + "a.MOT_RECH_5,a.COMENT_1_MOTRECH_5,a.COMENT_2_MOTRECH_5,"
                    + "a.COMENT_3_MOTRECH_5,a.MOT_RECH_6,a.COMENT_1_MOTRECH_6,"
                    + "a.COMENT_2_MOTRECH_6,a.COMENT_3_MOTRECH_6,a.MOT_RECH_7,"
                    + "a.COMENT_1_MOTRECH_7,a.COMENT_2_MOTRECH_7,a.COMENT_3_MOTRECH_7,"
                    + "a.MOT_RECH_8,a.COMENT_1_MOTRECH_8,a.COMENT_2_MOTRECH_8,"
                    + "a.COMENT_3_MOTRECH_8,a.MOT_RECH_9,a.COMENT_1_MOTRECH_9,"
                    + "a.COMENT_2_MOTRECH_9,a.COMENT_3_MOTRECH_9,a.MOT_RECH_10,"
                    + "a.COMENT_1_MOTRECH_10,a.COMENT_2_MOTRECH_10,a.COMENT_3_MOTRECH_10,"
                    + "a.TESTCONT,a.MANUAL,a.MED_BAJ,"
                    + "a.PT_CONT , a.PRC_PERD, a.MOD_FECHA_SOLCT, a.FECHA_SOLCT, a.MOD_FECHA_RESP, a.FECHA_RESP "
                    + "from SOLATRMT_ANT a Inner JOIN"
                    + "("
                    + "select max(fRecepcion) fRecepcion  , ccounips "
                    + "from SOLATRMT_ANT "
                    + "group by ccounips "
                    + ") AS B "
                    + "on a.ccounips= B.ccounips "
                    + "and a.fRecepcion= B.fRecepcion "
                    + "group by ccounips ";

                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();



            }
            catch (Exception e)
            {
                ficheroLog.AddError("SolicitudesATR.CargaSolAtr: " + e.Message);
            }

        }
        private void Genera_SOLATRMT_MAX()
        {

            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                strSql = "DELETE FROM SOLATRMT_MAX;";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "replace into SOLATRMT_MAX select a.EMP_TIT,a.estadoSolATR, a.codSolATR,"
                    + "a.tipoSolATR,a.fRecepcion,a.fAcepRech,"
                    + "a.fRechazo,a.fEnvioATR,a.fEnvioDoc,"
                    + "a.fPrevAlta,a.fActivacion,a.CCOUNIPS,"
                    + "a.CUPS_EXT,a.TIPO_CLI,a.CIF,"
                    + "a.cliente,a.SEG_MER,a.LINEA_N,"
                    + "a.DIREC_PS,a.CDISTRIB,a.CONTR_ATR,"
                    + "a.COMENT_1_SOLICITUD,a.COMENT_2_SOLICITUD,a.COMENT_3_SOLICITUD,"
                    + "a.CHECK_ANULAC,a.MOT_BAJA,a.POTENCIA1,"
                    + "a.POTENCIA2,a.POTENCIA3,a.POTENCIA4,"
                    + "a.POTENCIA5,a.POTENCIA6,a.TARIFA,"
                    + "a.TENSION,a.CCONTRPS,a.VER_CONTR_PS,"
                    + "a.USO_CONTR,a.GESTOR_SCE,a.COMENT_1_PETTRAB,"
                    + "a.COMENT_2_PETTRAB,a.COMENT_3_PETTRAB,a.MOT_RECH_1,"
                    + "a.COMENT_1_MOTRECH_1,a.COMENT_2_MOTRECH_1,a.COMENT_3_MOTRECH_1,"
                    + "a.MOT_RECH_2,a.COMENT_1_MOTRECH_2,a.COMENT_2_MOTRECH_2,"
                    + "a.COMENT_3_MOTRECH_2,a.MOT_RECH_3,a.COMENT_1_MOTRECH_3,"
                    + "a.COMENT_2_MOTRECH_3,a.COMENT_3_MOTRECH_3,a.MOT_RECH_4,"
                    + "a.COMENT_1_MOTRECH_4,a.COMENT_2_MOTRECH_4,a.COMENT_3_MOTRECH_4,"
                    + "a.MOT_RECH_5,a.COMENT_1_MOTRECH_5,a.COMENT_2_MOTRECH_5,"
                    + "a.COMENT_3_MOTRECH_5,a.MOT_RECH_6,a.COMENT_1_MOTRECH_6,"
                    + "a.COMENT_2_MOTRECH_6,a.COMENT_3_MOTRECH_6,a.MOT_RECH_7,"
                    + "a.COMENT_1_MOTRECH_7,a.COMENT_2_MOTRECH_7,a.COMENT_3_MOTRECH_7,"
                    + "a.MOT_RECH_8,a.COMENT_1_MOTRECH_8,a.COMENT_2_MOTRECH_8,"
                    + "a.COMENT_3_MOTRECH_8,a.MOT_RECH_9,a.COMENT_1_MOTRECH_9,"
                    + "a.COMENT_2_MOTRECH_9,a.COMENT_3_MOTRECH_9,a.MOT_RECH_10,"
                    + "a.COMENT_1_MOTRECH_10,a.COMENT_2_MOTRECH_10,a.COMENT_3_MOTRECH_10,"
                    + "a.TESTCONT,a.MANUAL,a.MED_BAJ,"
                    + "a.PT_CONT , a.PRC_PERD, a.MOD_FECHA_SOLCT, a.FECHA_SOLCT, a.MOD_FECHA_RESP, a.FECHA_RESP "
                    + "from SOLATRMT a Inner JOIN"
                    + "("
                    + "select ccounips, max(codSolATR) as maxCodSolATR "
                    + "from SOLATRMT "
                    + "group by ccounips "
                    + ") AS B "
                    + "on a.ccounips= B.ccounips "
                    + "and a.codSolATR= B.maxCodSolATR "
                    + "group by ccounips ";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch(Exception e)
            {
                ficheroLog.AddError(e.Message);
            }
        }
        private void PublicaEnFTP(FileInfo archivo)
        {
            utilidades.UltimateFTP ftp;
            try
            {
                ftp = new utilidades.UltimateFTP(
                        param.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                        param.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                        utilidades.FuncionesTexto.Decrypt(param.GetValue("ftp_pass", DateTime.Now, DateTime.Now), true),
                        param.GetValue("ftp_port", DateTime.Now, DateTime.Now));

                ftp.Upload(archivo.Name, archivo.FullName);

                NotificaEntregaSOLATRMT(archivo, "FTP");

            }
            catch(Exception e)
            {
                ficheroLog.AddError("SolicitudesATR.PublicaEnFTP: " + e.Message);
            }
        }

        private void NotificaEntregaSOLATRMT(FileInfo archivo, string sitio)
        {

            try
            {

                string from = param.GetValue("Buzon_envio_email");
                string to = param.GetValue("mail_para_aviso_ftp");
                //                string cc = param.GetValue("email_agora_copia");
                string subject = "Archivo " + archivo.Name + " entregado en FTP.";

                //mail.MailExchangeServer email = new mail.MailExchangeServer(System.Environment.UserName);
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
                StringBuilder textBody = new StringBuilder();

                textBody.Clear();
                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Ya se ha actualizado en el " + sitio + " el archivo ").Append(archivo.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");

                if (param.GetValue("EnviarMail") == "S")
                    mes.SendMail(from, to, null, subject, textBody.ToString(), null);

                else
                    mes.SaveMail(from, to, null, subject, textBody.ToString(), null);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("Buzon_envio_email"));


            }
            catch(Exception e)
            {
                ficheroLog.AddError("NotificaEntregaSOLATRMT: " + e.Message);
            }

           


           

        }

        private void Covid_Aplazamiento_Facturacion()
        {
            EndesaBusiness.contratacion.Covid covid_contratacion = new EndesaBusiness.contratacion.Covid();
            covid_contratacion.AplazadosQuePidenBaja();
        }

        public bool ExisteCUPS(string cups)
        {
            bool existe = false;
            EndesaEntity.contratacion.SolATRMT o;
            if(dic.TryGetValue(cups, out o))
            {
                existe = true;
                this.fAcepRech = o.fAcepRech;
                this.cups_ext = o.cups_ext;
            }
            return existe;
        }

    }
}
