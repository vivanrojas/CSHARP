using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class InformeComparativaConversionContratos
    {

        Dictionary<string, string> dic_relacion_complementos;
        Dictionary<string, string> dic_relacion_complementos_nuevos;
        Dictionary<string, string> dic_relacion_complementos_invertidos;
        Dictionary<string, EndesaEntity.facturacion.InformeComparativaContratos> dic_foto1;
        Dictionary<string, EndesaEntity.facturacion.InformeComparativaContratos> dic_foto2;
        
        utilidades.Param param;
        logs.Log ficheroLog;

        public InformeComparativaConversionContratos()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Revision_facturas");
            param = new utilidades.Param("global_param", servidores.MySQLDB.Esquemas.AUX);
            dic_relacion_complementos = Carga_Relacion_Complementos();
            dic_relacion_complementos_nuevos = Carga_Relacion_Complementos_Nuevos();
            dic_relacion_complementos_invertidos = Carga_Relacion_Complementos_Inversa();          

        }
                

        public void DescargaFicheroExtractor()
        {
            EndesaBusiness.utilidades.Fechas f = new utilidades.Fechas();
            string md5 = "";

            try
            {

                string archivo_origen = param.GetValue("COMP_CONT_PREF_ARCHI1_RUTA")
                    + param.GetValue("COMP_CONT_PREF_FICHERO_EXTRACTOR");


                string archivo_destino = param.GetValue("COMP_CONT_PREF_ARCHI1_RUTA")
                    + param.GetValue("COMP_CONT_PREF_ARCH1")
                    + f.UltimoDiaHabil().ToString("MMdd")
                    + param.GetValue("COMP_CONT_PREF_ARCH1_EXT");

                FileInfo file = new FileInfo(archivo_origen);
                FileInfo file_destino = new FileInfo(archivo_destino);

                ficheroLog.Add("Llamando a extractor --> " +
                        param.GetValue("COMP_CONT_PREF_EXTRACTOR") + " para el dia " +
                        utilidades.Fichero.UltimoDiaHabil_YYMMDD());
                utilidades.Fichero.EjecutaComando(param.GetValue("COMP_CONT_PREF_EXTRACTOR"), null);
                ficheroLog.Add("Fin extractor");

                if (file_destino.Exists)
                    file_destino.Delete();


                if (file.Exists)
                {
                    md5 = utilidades.Fichero.checkMD5(file.FullName).ToString();
                    if (file.Length > 0 && (md5 != param.GetValue("COMP_CONT_PREF_MD5_contratosPS")))
                    {
                        System.IO.File.Move(archivo_origen, archivo_destino);
                        param.code = "COMP_CONT_PREF_MD5_contratosPS";
                        param.from_date = new DateTime(2021, 06, 10);
                        param.to_date = new DateTime(4999, 12, 31);
                        param.value = md5;
                        param.Save();

                    }
                    else
                        file.Delete();
                        
                }
                    

            }catch(Exception e)
            {
                ficheroLog.AddError("DescargaFicheroExtractor: " + e.Message);
            }

        }

        public void Proceso(string ccounips)
        {
            string archivo_informe = "";
            EndesaBusiness.utilidades.Fechas f = new utilidades.Fechas();
            EndesaBusiness.utilidades.FechasProcesos fp = new utilidades.FechasProcesos();

            DescargaFicheroExtractor();


            Console.WriteLine("Última ejecución del proceso: "
                + fp.GetFechaProceso("INF_COMP_CONV_CONT").Date.ToString("dd/MM/yyyy")
                + " < " + DateTime.Now.Date.ToString("dd/MM/yyyy"));

            if (fp.GetFechaProceso("INF_COMP_CONV_CONT").Date < DateTime.Now.Date)
            {
                              


                string archivo = param.GetValue("COMP_CONT_PREF_ARCHI1_RUTA")
                + param.GetValue("COMP_CONT_PREF_ARCH1")
                + f.UltimoDiaHabil().ToString("MMdd")
                + param.GetValue("COMP_CONT_PREF_ARCH1_EXT");

                FileInfo fileInfo = new FileInfo(archivo);
                if (fileInfo.Exists)
                {

                    ImportarArchivo(archivo, "irf_contratos_foto2");

                    dic_foto1 = CargaFotoContratos("irf_contratos_foto1", ccounips);
                    dic_foto2 = CargaFotoContratos("irf_contratos_foto2", ccounips);
                    Analiza();
                    archivo_informe = GeneraInforme();
                    EnvioCorreo(archivo_informe);
                    fp.Update("INF_COMP_CONV_CONT", DateTime.Now);
                    fileInfo.Delete();

                }
                else
                {
                    EnvioCorreoNoExisteArchivo(archivo);
                }

                
            }

            
        }

        public void GeneraSoloInforme()
        {
            dic_foto1 = CargaFotoContratos("irf_contratos_foto1", null);
            dic_foto2 = CargaFotoContratos("irf_contratos_foto2", null);
            Analiza();
            GeneraInforme();
        }


        private void EnvioCorreo(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                string from = param.GetValue("COMP_CONT_PREF_BUZON_ENVIO");
                string to = param.GetValue("fac_inf_conv_cont_email_para");
                // string cc = param.GetValue("fac_inf_conv_cont_email_cc");
                string subject = param.GetValue("fac_inf_conv_cont_mail_asunto") + " a " + DateTime.Now.ToString("dd/MM/yyyy");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Se adjunta el archivo ").Append(fileInfo.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");

                // EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();


                if (param.GetValue("fac_inf_conv_cont_enviar_mail") == "S")
                    mes.SendMail(from, to, null, subject, textBody.ToString(), archivo);

                else
                    mes.SaveMail(from, to, null, subject, textBody.ToString(), archivo);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("COMP_CONT_PREF_BUZON_ENVIO"));
                fileInfo.Delete();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }
        private void EnvioCorreoNoExisteArchivo(string archivo)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                string from = param.GetValue("COMP_CONT_PREF_BUZON_ENVIO");
                string to = param.GetValue("fac_inf_conv_cont_email_falta_archivo_para");
                //string cc = param.GetValue("fac_inf_conv_cont_email_falta_archivo_cc");
                string subject = param.GetValue("fac_inf_conv_cont_mail_falta_asunto");

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  No se ha encontrado el archivo ").Append(fileInfo.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Sin este archivo de extracción no se puede generar el informe. ");
                
                textBody.Append("  Dentro de 15 minutos se volverá a comprobar su existencia. ");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");

                //EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();

                if (param.GetValue("fac_inf_conv_cont_enviar_mail") == "S")
                    mes.SendMail(from, to, null, subject, textBody.ToString(), null);

                else
                    mes.SaveMail(from, to, null, subject, textBody.ToString(), null);

                ficheroLog.Add("Correo enviado desde: " + param.GetValue("COMP_CONT_PREF_BUZON_ENVIO"));
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }

        private void ImportarArchivo(string archivo, string tabla)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            string line = "";
            string[] campos;
            int i = 0;
            int c = 0;
            string strSql = "";
            int total_registros = 0;


            try
            {
                
                strSql = "delete from " + tabla;
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                ficheroLog.Add("Importando archivo " + archivo);

                System.IO.StreamReader file = new System.IO.StreamReader(archivo, System.Text.Encoding.GetEncoding(1252));
                while ((line = file.ReadLine()) != null)
                {
                    campos = line.Split(';');
                    i++;
                    c = 1;
                    total_registros++;

                    if (firstOnly)
                    {
                        sb.Append("replace into ").Append(tabla);
                        sb.Append(" (CCOUNIPS, CONTREXT, TESTCONT, FALTACON, FPSERCON, FPREVBAJ, FBAJACON,");
                        sb.Append(" CEMPTITU, CCONTRPS, CNUMSCPS, CLINNEG, CCLIENTE, CCALENPO, VDIAFACT, FSIGFACT,");
                        sb.Append(" FFINVESU, CTARIFA, CUPSREE, TPUNTMED, CCOMPOBJ, VNSEGHOR, VNUMTRAM, VPARAM01,");
                        sb.Append(" VPARAM02, VPARAM03, VPARAM04, VPARAM05, CCOMAUTO, TTICONPS, TPOTENCIP1, VPOTCALIP1,");
                        sb.Append(" TPOTENCIP2, VPOTCALIP2, TPOTENCIP3, VPOTCALIP3, TPOTENCIP4, VPOTCALIP4, TPOTENCIP5,");
                        sb.Append(" VPOTCALIP5, TPOTENCIP6, VPOTCALIP6) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(campos[c]).Append("',"); c++;
                    sb.Append("'").Append(campos[c]).Append("',"); c++;

                    for(int j = 1; j <= 6; j++)
                    {
                        sb.Append(CF(campos[c])).Append(","); c++;
                    }

                    sb.Append(CN(campos[c])).Append(","); c++; // CCONTRPS

                    for (int j = 1; j <= 7; j++)
                    {
                        sb.Append(CF(campos[c])).Append(","); c++;
                    }

                    sb.Append(CN(campos[c])).Append(","); c++; // CTARIFA
                    sb.Append(CN(campos[c])).Append(","); c++; // CUPSREE
                    sb.Append(CF(campos[c])).Append(","); c++; // TPUNTMED
                    sb.Append(CN(campos[c])).Append(","); c++; // CCOMPOBJ
                    sb.Append(CF(campos[c])).Append(","); c++; // VNSEGHOR
                    sb.Append(CF(campos[c])).Append(","); c++; // VNUMTRAM

                    for (int j = 1; j <= 5; j++)
                    {
                        sb.Append(CDouble(campos[c])).Append(","); c++;
                    }

                    sb.Append(CF(campos[c])).Append(","); c++; // CCOMAUTO
                    sb.Append(CN(campos[c])); c++; // TTICONPS

                    for (int j = 1; j <= 6; j++)
                    {
                        sb.Append(",").Append(CN(campos[c])); c++; // TPOTENCIP1
                        sb.Append(",").Append(CDouble(campos[c])); c++; // VPOTCALIP1
                    }

                    sb.Append("),");


                    if (i == 250)
                    {
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
                file.Close();

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

                ficheroLog.Add("Fin importación");
                ficheroLog.Add("Se han importado " + total_registros + " registros");
                FileInfo archivo_info = new FileInfo(archivo);
                archivo_info.Delete();

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private Dictionary<string, EndesaEntity.facturacion.InformeComparativaContratos> CargaFotoContratos(string tabla, string ccounips)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string tarifa = "";
            int faltacon = 0;
            int fpsercon = 0;
            int fprevbaj = 0;
            int fbajacon = 0;

            Dictionary<string, EndesaEntity.facturacion.InformeComparativaContratos> d
                = new Dictionary<string, EndesaEntity.facturacion.InformeComparativaContratos>();

            try
            {
                              


                strSql = "SELECT t.CCOUNIPS, t.CONTREXT, t.TESTCONT, t.FALTACON, t.FPSERCON, t.FPREVBAJ, t.FBAJACON,"
                    + " t.CEMPTITU, t.CCONTRPS, t.CNUMSCPS, t.CLINNEG, t.CCLIENTE, t.CCALENPO, t.VDIAFACT, t.FSIGFACT,"
                    + " t.FFINVESU, t.CTARIFA, t.CUPSREE, t.TPUNTMED, t.CCOMPOBJ, t.VNSEGHOR, t.VNUMTRAM, t.VPARAM01,"
                    + " t.VPARAM02, t.VPARAM03, t.VPARAM04, t.VPARAM05, t.CCOMAUTO, t.TTICONPS, t.TPOTENCIP1, t.VPOTCALIP1,"
                    + " t.TPOTENCIP2, t.VPOTCALIP2, t.TPOTENCIP3, t.VPOTCALIP3, t.TPOTENCIP4, t.VPOTCALIP4, t.TPOTENCIP5,"
                    + " t.VPOTCALIP5, t.TPOTENCIP6, t.VPOTCALIP6, ps.TENSION"
                    + " FROM cont." + tabla + " t"
                    + " left outer join PS_AT ps on"
                    + " ps.IDU = t.CCOUNIPS and"
                    + " ps.CONTREXT = t.CONTREXT";

                if (ccounips != null)
                    strSql += " where CCOUNIPS = '" + ccounips + "'";
                    // + " WHERE TESTCONT in (3)";



                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {  

                    EndesaEntity.facturacion.InformeComparativaContratos o;
                    if (!d.TryGetValue(r["CONTREXT"].ToString(), out o))
                    {

                        tarifa = r["CTARIFA"].ToString();

                        if (r["FALTACON"] != System.DBNull.Value)
                            faltacon = Convert.ToInt32(r["FALTACON"]);

                        if (r["FPSERCON"] != System.DBNull.Value)
                            fpsercon = Convert.ToInt32(r["FPSERCON"]);

                        if (r["FPREVBAJ"] != System.DBNull.Value)
                            fprevbaj = Convert.ToInt32(r["FPREVBAJ"]);

                        if (r["FBAJACON"] != System.DBNull.Value)
                            fbajacon = Convert.ToInt32(r["FBAJACON"]);
                        
                        o = new EndesaEntity.facturacion.InformeComparativaContratos();
                        o.cups = r["CCOUNIPS"].ToString();
                        o.contrato = r["CONTREXT"].ToString();
                        o.version_cierre = Convert.ToInt32(r["CNUMSCPS"]);

                        if (r["TENSION"] != System.DBNull.Value)
                            o.tension = Convert.ToInt32(r["TENSION"]);

                        o.dic = Cabecera(tarifa, faltacon, fpsercon, fprevbaj, fbajacon);

                        EndesaEntity.facturacion.InformeComparativaContratosDetalle oo
                           = new EndesaEntity.facturacion.InformeComparativaContratosDetalle();
                                                
                        oo.valor_antes = r["CCOMPOBJ"].ToString();
                        oo.contraste = "KO";
                        o.dic.Add(oo.valor_antes, oo);
                        d.Add(o.contrato, o);

                    }
                    else
                    {
                        EndesaEntity.facturacion.InformeComparativaContratosDetalle oo
                         = new EndesaEntity.facturacion.InformeComparativaContratosDetalle();

                        oo.valor_antes = r["CCOMPOBJ"].ToString();
                        oo.contraste = "KO";
                        o.dic.Add(oo.valor_antes, oo);
                    }
                                           

                }
                db.CloseConnection();
            
                return d;
            }catch(Exception e)
            {
                return null;
            }
        }

        private Dictionary<string, EndesaEntity.facturacion.InformeComparativaContratosDetalle> 
            Cabecera(string tarifa, int faltacon, int fpsercon, int fprevbaj, int fbajacon)
        {

            Dictionary<string, EndesaEntity.facturacion.InformeComparativaContratosDetalle> d
                = new Dictionary<string, EndesaEntity.facturacion.InformeComparativaContratosDetalle>();

            EndesaEntity.facturacion.InformeComparativaContratosDetalle c = new EndesaEntity.facturacion.InformeComparativaContratosDetalle();
            c.valor_antes = tarifa;
            c.contraste = "KO";
            d.Add("TARIFA", c);
                       
            c = new EndesaEntity.facturacion.InformeComparativaContratosDetalle();
            c.valor_antes = faltacon > 0 ? faltacon.ToString() : "";
            c.contraste = "KO";
            d.Add("FALTACON", c);
           
            c = new EndesaEntity.facturacion.InformeComparativaContratosDetalle();
            c.valor_antes = fpsercon > 0 ? fpsercon.ToString() : "";
            c.contraste = "KO";
            d.Add("FPSERCON", c);
           
            c = new EndesaEntity.facturacion.InformeComparativaContratosDetalle();
            c.valor_antes = fprevbaj > 0 ? fprevbaj.ToString() : "";
            c.contraste = "KO";
            d.Add("FPREVBAJ", c);           

            c = new EndesaEntity.facturacion.InformeComparativaContratosDetalle();
            c.valor_antes = fbajacon > 0 ? fbajacon.ToString() : "";
            c.contraste = "KO";
            d.Add("FBAJACON", c);

            return d;

        }
                

        public void Analiza()
        {
            // Utilizamos foto1 como base
            // Buscamos los valores para complementar en foto2
            // Complementamos el valor ahora con lo encontrado en foto2

            EndesaEntity.facturacion.InformeComparativaContratosDetalle oo;

            foreach (KeyValuePair<string, EndesaEntity.facturacion.InformeComparativaContratos> p in dic_foto1)
            {                             

                EndesaEntity.facturacion.InformeComparativaContratos foto2;
                // Buscamos el CONTRATO EXTERNO (KEY)
                if (dic_foto2.TryGetValue(p.Key, out foto2))
                {
                    p.Value.version_conversion = foto2.version_cierre;
                    // Encontramos foto2
                    foreach (KeyValuePair<string, EndesaEntity.facturacion.InformeComparativaContratosDetalle> pp in p.Value.dic)
                    {

                        switch (pp.Key)
                        {
                            case "TARIFA":
                                if (foto2.dic.TryGetValue(pp.Key, out oo))
                                {
                                    pp.Value.valor_ahora = oo.valor_antes;
                                    pp.Value.contraste = CompruebaTarifa(pp.Value.valor_antes, oo.valor_antes, p.Value.tension);
                                }
                                
                                break;
                            case "FPSERCON":
                                if (foto2.dic.TryGetValue(pp.Key, out oo))
                                {
                                    pp.Value.valor_ahora = oo.valor_antes;
                                    if (oo.valor_antes == "20210601")
                                        pp.Value.contraste = "OK";
                                }
                                break;
                            case "FALTACON":
                                if (foto2.dic.TryGetValue(pp.Key, out oo))
                                {
                                    pp.Value.valor_ahora = oo.valor_antes;
                                    pp.Value.contraste = (pp.Value.valor_antes == oo.valor_antes ? "OK" : "KO");
                                }                                    
                                break;
                            case "FPREVBAJ":
                                if (foto2.dic.TryGetValue(pp.Key, out oo))
                                {
                                    pp.Value.valor_ahora = oo.valor_antes;
                                    pp.Value.contraste = (pp.Value.valor_antes == oo.valor_antes ? "OK" : "KO");
                                }
                                break;
                            case "FBAJACON":
                                if (foto2.dic.TryGetValue(pp.Key, out oo))
                                {
                                    pp.Value.valor_ahora = oo.valor_antes;
                                    pp.Value.contraste = (pp.Value.valor_antes == oo.valor_antes ? "OK" : "KO");
                                }
                                break;
                            default:
                                if (foto2.dic.TryGetValue(Posterior_Circular(pp.Value.valor_antes), out oo))
                                {                                    
                                    pp.Value.valor_ahora = oo.valor_antes;
                                    pp.Value.contraste = "OK";                                    
                                }
                                break;

                        }

                    }
                }                   
            }

            foreach (KeyValuePair<string, EndesaEntity.facturacion.InformeComparativaContratos> p in dic_foto2)
            {
                EndesaEntity.facturacion.InformeComparativaContratos foto1;
                // Buscamos el CONTRATO EXTERNO (KEY)
                if (dic_foto1.TryGetValue(p.Key, out foto1))
                {
                    // Encontramos foto1
                    foreach (KeyValuePair<string, EndesaEntity.facturacion.InformeComparativaContratosDetalle> pp in p.Value.dic)
                    {
                        string complemento_nuevo;
                        if (dic_relacion_complementos_nuevos.TryGetValue(pp.Key, out complemento_nuevo))
                        {
                            EndesaEntity.facturacion.InformeComparativaContratosDetalle c
                                = new EndesaEntity.facturacion.InformeComparativaContratosDetalle();
                            c.contraste = "OK";
                            c.valor_antes = "";
                            c.valor_ahora = complemento_nuevo;
                            foto1.dic.Add(complemento_nuevo, c);
                        }
                        else if(dic_relacion_complementos_invertidos.TryGetValue(pp.Key, out complemento_nuevo))
                            if(!foto1.dic.TryGetValue(complemento_nuevo, out oo))
                            {
                                EndesaEntity.facturacion.InformeComparativaContratosDetalle c
                                = new EndesaEntity.facturacion.InformeComparativaContratosDetalle();

                                EndesaEntity.facturacion.InformeComparativaContratosDetalle tarifa;
                                foto1.dic.TryGetValue("TARIFA", out tarifa);
                                
                                if(pp.Key == "L86" && tarifa.valor_antes == "3.1A")
                                {
                                    c.contraste = "OK";
                                    c.valor_antes = "";
                                    c.valor_ahora = pp.Key;
                                    foto1.dic.Add(pp.Key, c);
                                }
                                else
                                {
                                    c.contraste = "KO";
                                    c.valor_antes = "";
                                    c.valor_ahora = pp.Key;
                                    foto1.dic.Add(pp.Key, c);
                                }
                              
                            }                           
                        
                        
                    }

                }
            }


        }


        private string CompruebaTarifa(string antes, string ahora, int tension)
        {

            if(tension > 0)
            {
                if (antes.Substring(0, 3) == ahora.Substring(0, 3) && ahora.Contains("TD"))
                    return "OK";
                else if (antes == "3.1A")
                {
                    if (tension < 30000 && ahora == "6.1TD")
                        return "OK";
                    else if (tension >= 30000 && ahora == "6.2TD")
                        return "OK";
                }

            }
            else
            {
                return ahora.Contains("TD") ? "OK" : "KO";
            }
           

            return "KO";

        }

        private Dictionary<string, string> Carga_Relacion_Complementos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, string> d
                = new Dictionary<string, string>();

            try
            {
                strSql = "select previo_circular, posterior_circular"
                    + " from fact_p_relacion_complementos where"
                    + " previo_circular is not null";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    d.Add(r["previo_circular"].ToString(), r["posterior_circular"].ToString());
                }
                db.CloseConnection();
                return d;
            }
            catch(Exception e)
            {
                return null;
            }
        }

        private Dictionary<string, string> Carga_Relacion_Complementos_Inversa()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, string> d
                = new Dictionary<string, string>();

            try
            {
                strSql = "select previo_circular, posterior_circular"
                    + " from fact_p_relacion_complementos where"
                    + " previo_circular is not null group by posterior_circular";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {                    
                    d.Add(r["posterior_circular"].ToString(), r["previo_circular"].ToString());
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        private Dictionary<string, string> Carga_Relacion_Complementos_Nuevos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, string> d
                = new Dictionary<string, string>();

            try
            {
                strSql = "select previo_circular, posterior_circular"
                    + " from fact_p_relacion_complementos where"
                    + " previo_circular is null";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    d.Add(r["posterior_circular"].ToString(), r["posterior_circular"].ToString());
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        


        private string Posterior_Circular(string previo_circular)
        {
            string o;

            if (param.GetValue("utilizar_tabla_conversion") != "N")
            {
                if (dic_relacion_complementos.TryGetValue(previo_circular, out o))
                    return o;
                else
                    return previo_circular;
            }
            else
                return previo_circular;
        }

        public string GeneraInforme()
        {

            ExcelPackage excelPackage;
            FileInfo fileInfo;
            string variable = "";

            int f = 1;
            int c = 1;

            try
            {

                fileInfo = new FileInfo(param.GetValue("fac_inf_conv_cont_ruta_salida")
                + param.GetValue("fac_inf_conv_cont_nombre_informe")  
                + DateTime.Now.ToString("yyyy_MM_dd_HHmmss") + ".xlsx");

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("Comparativa");

                var headerCells = workSheet.Cells[1, 1, 1, 8];
                var headerFont = headerCells.Style.Font;

                var allCells = workSheet.Cells[1, 1, 1, 8];
                var cellFont = allCells.Style.Font;
                cellFont.Bold = true;

                workSheet.Cells[f, c].Value = "CUPS";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;
                
                workSheet.Cells[f, c].Value = "CONTRATO";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "VERSION CIERRE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "VERSION CONVERSION";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "VARIABLE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;               

                workSheet.Cells[f, c].Value = "FOTO DEL CIERRE DEL MES";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;
               
                workSheet.Cells[f, c].Value = "FOTO TRAS LA CONVERSIÓN";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "CONTRASTE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                /*
                d.Add("TARIFA", c);

                c = new EndesaEntity.facturacion.InformeComparativaContratosDetalle();
                c.valor_antes = faltacon > 0 ? faltacon.ToString() : "";
                d.Add("FALTACON", c);

                c = new EndesaEntity.facturacion.InformeComparativaContratosDetalle();
                c.valor_antes = fpsercon > 0 ? fpsercon.ToString() : "";
                d.Add("FPSERCON", c);

                c = new EndesaEntity.facturacion.InformeComparativaContratosDetalle();
                c.valor_antes = fprevbaj > 0 ? fprevbaj.ToString() : "";
                d.Add("FPREVBAJ", c);

                c = new EndesaEntity.facturacion.InformeComparativaContratosDetalle();
                c.valor_antes = fbajacon > 0 ? fbajacon.ToString() : "";
                d.Add("FBAJACON", c);
                */

                foreach (KeyValuePair<string, EndesaEntity.facturacion.InformeComparativaContratos> p in dic_foto1)
                {                 
                                       
                    foreach (KeyValuePair<string, EndesaEntity.facturacion.InformeComparativaContratosDetalle> pp in p.Value.dic)
                    {
                        c = 1;
                        f++;

                        workSheet.Cells[f, c].Value = p.Value.cups; c++;
                        workSheet.Cells[f, c].Value = p.Value.contrato; c++;
                        workSheet.Cells[f, c].Value = p.Value.version_cierre; c++;
                        workSheet.Cells[f, c].Value = p.Value.version_conversion; c++;

                        switch (pp.Key)
                        {
                            case "TARIFA":
                                variable = pp.Key;
                                break;
                            case "FALTACON":
                                variable = pp.Key;
                                break;
                            case "FPSERCON":
                                variable = pp.Key;                               
                                break;
                            case "FPREVBAJ":
                                variable = pp.Key;
                                break;
                            case "FBAJACON":
                                variable = pp.Key;
                                break;
                            default:
                                variable = "Complemento";
                                break;

                        }
                        workSheet.Cells[f, c].Value = variable; c++;
                        workSheet.Cells[f, c].Value = pp.Value.valor_antes; c++;
                        workSheet.Cells[f, c].Value = pp.Value.valor_ahora; c++;
                        workSheet.Cells[f, c].Value = pp.Value.contraste; c++;

                    }
                    
                }


                allCells = workSheet.Cells[1, 1, f, 8];                

                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:H1"].AutoFilter = true;
                allCells.AutoFitColumns();

                excelPackage.Save();

                return fileInfo.FullName;

            }
            catch(Exception e)
            {                
                Console.WriteLine(e.Message);
                return null;
            }
        }


        private string CF(string t)
        {
            if (t.Trim() == "00000000" || t.Trim() == "")
                return "null";
            else
                return t.Trim();
        }

        private string CN(string t)
        {
            if (t.Trim() == "00000000" || t.Trim() == "")
                return "'null'";
            else
                return "'" + t.Trim() + "'";
        }

        public string CDouble(String t)
        {
            t = t.Trim();
            if (t == "")
            {
                return "null";
            }
            else
            {
                t = t.Replace(" ", "");
                t = t.Replace("+", string.Empty);
                t = t.Replace("----------", string.Empty);
                t = t.Replace(".", string.Empty);
                t = t.Replace(",", ".");

                if (t == "")
                {
                    t = "null";
                }
            }

            return t;
        }

    }
}


