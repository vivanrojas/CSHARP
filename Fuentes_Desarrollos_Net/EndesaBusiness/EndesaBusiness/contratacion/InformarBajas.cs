using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion
{
    public class InformarBajas
    {

        logs.Log ficheroLog;
        utilidades.Param param;
        public InformarBajas()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_InformarBajas");
            param = new utilidades.Param("informar_bajas_param", servidores.MySQLDB.Esquemas.CON);
        }

        public void CreaInformeATR()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            int f = 0;
            int c = 0;

            List<string> lista_territorios = new List<string>();

            // donde string --> responsable territorial
            Dictionary<string, List<EndesaEntity.contratacion.SOLATRMT_InformeATR_Carterizado>> dic =
                new Dictionary<string, List<EndesaEntity.contratacion.SOLATRMT_InformeATR_Carterizado>>();

            try
            {

                // Solicitudes del último mes.

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
                    + " WHERE sol.fRecepcion > date_format((now() + interval - (1) month),'%Y%m%d') and"
                    + " cartera.responsable_territorial <> ''";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.SOLATRMT_InformeATR_Carterizado cc =
                        new EndesaEntity.contratacion.SOLATRMT_InformeATR_Carterizado();

                    #region campos

                    if (r["EMP_TIT"] != System.DBNull.Value)
                        cc.empresa_titular = Convert.ToInt32(r["EMP_TIT"]);

                    if (r["estadoSolATR"] != System.DBNull.Value)
                        cc.estadoSolAtr = r["estadoSolATR"].ToString();

                    if (r["codSolATR"] != System.DBNull.Value)
                        cc.codSolAtr = Convert.ToInt64(r["codSolATR"]);

                    if (r["tipoSolATR"] != System.DBNull.Value)
                        cc.tipoSolAtr = r["tipoSolATR"].ToString();

                    if (r["fRecepcion"] != System.DBNull.Value)
                        cc.fRecepcion = Convert.ToInt32(r["fRecepcion"]);

                    if (r["fAcepRech"] != System.DBNull.Value)
                        cc.fAcepRech = Convert.ToInt32(r["fAcepRech"]);

                    if (r["fRechazo"] != System.DBNull.Value)
                        cc.fRechazo = Convert.ToInt32(r["fRechazo"]);

                    if (r["fEnvioATR"] != System.DBNull.Value)
                        cc.fEnvioAtr = Convert.ToInt32(r["fEnvioATR"]);

                    if (r["fEnvioDoc"] != System.DBNull.Value)
                        cc.fEnvioDoc = Convert.ToInt32(r["fEnvioDoc"]);

                    if (r["fPrevAlta"] != System.DBNull.Value)
                        cc.fPrevAlta = Convert.ToInt32(r["fPrevAlta"]);

                    if (r["fActivacion"] != System.DBNull.Value)
                        cc.fActivacion = Convert.ToInt32(r["fActivacion"]);

                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        cc.ccounips = r["CCOUNIPS"].ToString();

                    if (r["CUPS_EXT"] != System.DBNull.Value)
                        cc.cups_ext = r["CUPS_EXT"].ToString();

                    if (r["TIPO_CLI"] != System.DBNull.Value)
                        cc.tipo_cli = r["TIPO_CLI"].ToString();

                    if (r["CIF"] != System.DBNull.Value)
                        cc.cif = r["CIF"].ToString();

                    if (r["cliente"] != System.DBNull.Value)
                        cc.cliente = r["cliente"].ToString();

                    if (r["SEG_MER"] != System.DBNull.Value)
                        cc.seg_mer = r["SEG_MER"].ToString();

                    if (r["LINEA_N"] != System.DBNull.Value)
                        cc.linea_n = Convert.ToInt32(r["LINEA_N"]);

                    if (r["DIREC_PS"] != System.DBNull.Value)
                        cc.direc_ps = r["DIREC_PS"].ToString();

                    if (r["CDISTRIB"] != System.DBNull.Value)
                        cc.cdistrib = r["CDISTRIB"].ToString();

                    if (r["CONTR_ATR"] != System.DBNull.Value)
                        cc.contr_atr = Convert.ToInt64(r["CONTR_ATR"]);

                    if (r["COMENT_1_SOLICITUD"] != System.DBNull.Value)
                        cc.coment_1_solicitud = r["COMENT_1_SOLICITUD"].ToString();

                    if (r["COMENT_2_SOLICITUD"] != System.DBNull.Value)
                        cc.coment_2_solicitud = r["COMENT_2_SOLICITUD"].ToString();

                    if (r["COMENT_3_SOLICITUD"] != System.DBNull.Value)
                        cc.coment_3_solicitud = r["COMENT_3_SOLICITUD"].ToString();

                    if (r["CHECK_ANULAC"] != System.DBNull.Value)
                        cc.check_anulac = r["CHECK_ANULAC"].ToString();

                    if (r["MOT_BAJA"] != System.DBNull.Value)
                        cc.mot_baja = r["MOT_BAJA"].ToString();

                    for(int i = 1; i <= 6; i++)                    
                        if (r["POTENCIA" + i] != System.DBNull.Value)
                            cc.potencia[i] = Convert.ToDouble(r["POTENCIA" + i]);

                    if (r["TARIFA"] != System.DBNull.Value)
                        cc.tarifa = r["TARIFA"].ToString();

                    if (r["TENSION"] != System.DBNull.Value)
                        cc.tension = Convert.ToInt32(r["TENSION"]);

                    if (r["CCONTRPS"] != System.DBNull.Value)
                        cc.ccontrps = r["CCONTRPS"].ToString();

                    if (r["VER_CONTR_PS"] != System.DBNull.Value)
                        cc.ver_contr_ps = Convert.ToInt32(r["VER_CONTR_PS"]);

                    if (r["USO_CONTR"] != System.DBNull.Value)
                        cc.uso_contr = r["USO_CONTR"].ToString();

                    if (r["GESTOR_SCE"] != System.DBNull.Value)
                        cc.gestor_sce = r["GESTOR_SCE"].ToString();

                    if (r["territorio"] != System.DBNull.Value)
                        cc.territorio = r["territorio"].ToString();

                    if (r["responsable_territorial"] != System.DBNull.Value)
                        cc.responsable_territorial = r["responsable_territorial"].ToString();

                    if (r["email_gestor"] != System.DBNull.Value)
                        cc.mail_responsable_territorial = r["email_gestor"].ToString();

                    #endregion

                    List<EndesaEntity.contratacion.SOLATRMT_InformeATR_Carterizado> o;
                    if (!dic.TryGetValue(cc.mail_responsable_territorial, out o))
                    {
                        o = new List<EndesaEntity.contratacion.SOLATRMT_InformeATR_Carterizado>();
                        o.Add(cc);
                        dic.Add(cc.mail_responsable_territorial, o);
                    }
                    else
                        o.Add(cc);

                }
                db.CloseConnection();

                foreach (KeyValuePair<string, List<EndesaEntity.contratacion.SOLATRMT_InformeATR_Carterizado>> p in dic)
                {
                    string salida_informes = param.GetValue("salida_informe");
                    FileInfo fileInfo = new FileInfo(salida_informes + "InformeATR_" 
                        + DateTime.Now.ToString("yyyy_MM_dd_HHmmss") + ".xlsx");
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                    ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                    var workSheet = excelPackage.Workbook.Worksheets.Add("InformeATR");

                    var headerCells = workSheet.Cells[1, 1, 1, 39];
                    var headerFont = headerCells.Style.Font;
                    f = 1;
                    c = 1;
                    headerFont.Bold = true;

                    workSheet.Cells[f, c].Value = "EMP_TIT"; c++;
                    workSheet.Cells[f, c].Value = "estadoSolATR"; c++;
                    workSheet.Cells[f, c].Value = "codSolATR"; c++;
                    workSheet.Cells[f, c].Value = "tipoSolATR"; c++;
                    workSheet.Cells[f, c].Value = "fRecepcion"; c++;
                    workSheet.Cells[f, c].Value = "fAcepRech"; c++;
                    workSheet.Cells[f, c].Value = "fRechazo"; c++;
                    workSheet.Cells[f, c].Value = "fEnvioATR"; c++;
                    workSheet.Cells[f, c].Value = "fEnvioDoc"; c++;
                    workSheet.Cells[f, c].Value = "fPrevAlta"; c++;
                    workSheet.Cells[f, c].Value = "fActivacion"; c++;
                    workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
                    workSheet.Cells[f, c].Value = "CUPS_EXT"; c++;
                    workSheet.Cells[f, c].Value = "TIPO_CLI"; c++;
                    workSheet.Cells[f, c].Value = "CIF"; c++;
                    workSheet.Cells[f, c].Value = "cliente"; c++;
                    workSheet.Cells[f, c].Value = "SEG_MER"; c++;
                    workSheet.Cells[f, c].Value = "LINEA_N"; c++;
                    workSheet.Cells[f, c].Value = "DIREC_PS"; c++;
                    workSheet.Cells[f, c].Value = "CDISTRIB"; c++;
                    workSheet.Cells[f, c].Value = "CONTR_ATR"; c++;
                    workSheet.Cells[f, c].Value = "COMENT_1_SOLICITUD"; c++;
                    workSheet.Cells[f, c].Value = "COMENT_2_SOLICITUD"; c++;
                    workSheet.Cells[f, c].Value = "COMENT_3_SOLICITUD"; c++;
                    workSheet.Cells[f, c].Value = "CHECK_ANULAC"; c++;
                    workSheet.Cells[f, c].Value = "MOT_BAJA"; c++;
                    workSheet.Cells[f, c].Value = "POTENCIA1"; c++;
                    workSheet.Cells[f, c].Value = "POTENCIA2"; c++;
                    workSheet.Cells[f, c].Value = "POTENCIA3"; c++;
                    workSheet.Cells[f, c].Value = "POTENCIA4"; c++;
                    workSheet.Cells[f, c].Value = "POTENCIA5"; c++;
                    workSheet.Cells[f, c].Value = "POTENCIA6"; c++;
                    workSheet.Cells[f, c].Value = "TARIFA"; c++;
                    workSheet.Cells[f, c].Value = "TENSION"; c++;
                    workSheet.Cells[f, c].Value = "CCONTRPS"; c++;
                    workSheet.Cells[f, c].Value = "VER_CONTR_PS"; c++;
                    workSheet.Cells[f, c].Value = "USO_CONTR"; c++;
                    workSheet.Cells[f, c].Value = "GESTOR_SCE"; c++;
                    workSheet.Cells[f, c].Value = "TERRITORIO"; c++;

                    foreach (EndesaEntity.contratacion.SOLATRMT_InformeATR_Carterizado pp in p.Value)
                    {
                        f++;
                        c = 1;

                        #region guardamos los territorios tratados
                        if (!lista_territorios.Exists(z => z == pp.territorio))
                            lista_territorios.Add(pp.territorio);
                        #endregion

                        workSheet.Cells[f, c].Value = pp.empresa_titular; c++;
                        workSheet.Cells[f, c].Value = pp.estadoSolAtr; c++;
                        workSheet.Cells[f, c].Value = pp.codSolAtr; c++;
                        workSheet.Cells[f, c].Value = pp.tipoSolAtr; c++;
                        workSheet.Cells[f, c].Value = pp.fRecepcion; c++;
                        workSheet.Cells[f, c].Value = pp.fAcepRech; c++;
                        workSheet.Cells[f, c].Value = pp.fRechazo; c++;
                        workSheet.Cells[f, c].Value = pp.fEnvioAtr; c++;
                        workSheet.Cells[f, c].Value = pp.fEnvioDoc; c++;
                        workSheet.Cells[f, c].Value = pp.fPrevAlta; c++;
                        workSheet.Cells[f, c].Value = pp.fActivacion; c++;
                        workSheet.Cells[f, c].Value = pp.ccounips; c++;
                        workSheet.Cells[f, c].Value = pp.cups_ext; c++;
                        workSheet.Cells[f, c].Value = pp.tipo_cli; c++;
                        workSheet.Cells[f, c].Value = pp.cif; c++;
                        workSheet.Cells[f, c].Value = pp.cliente; c++;
                        workSheet.Cells[f, c].Value = pp.seg_mer; c++;
                        workSheet.Cells[f, c].Value = pp.linea_n; c++;
                        workSheet.Cells[f, c].Value = pp.direc_ps; c++;
                        workSheet.Cells[f, c].Value = pp.cdistrib; c++;
                        workSheet.Cells[f, c].Value = pp.contr_atr; c++;
                        workSheet.Cells[f, c].Value = pp.coment_1_solicitud; c++;
                        workSheet.Cells[f, c].Value = pp.coment_2_solicitud; c++;
                        workSheet.Cells[f, c].Value = pp.coment_3_solicitud; c++;
                        workSheet.Cells[f, c].Value = pp.check_anulac; c++;
                        workSheet.Cells[f, c].Value = pp.mot_baja; c++;

                        for (int i = 1; i <= 6; i++)
                        {
                            workSheet.Cells[f, c].Value = pp.potencia[i];
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                        }


                        workSheet.Cells[f, c].Value = pp.tarifa; c++;

                        workSheet.Cells[f, c].Value = pp.tension;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                        workSheet.Cells[f, c].Value = pp.ccontrps; c++;
                        workSheet.Cells[f, c].Value = pp.ver_contr_ps; c++;
                        workSheet.Cells[f, c].Value = pp.uso_contr; c++;
                        workSheet.Cells[f, c].Value = pp.gestor_sce; c++;
                        workSheet.Cells[f, c].Value = pp.territorio; c++;
                    }

                    var allCells = workSheet.Cells[1, 1, f, c];
                    workSheet.View.FreezePanes(2, 1);
                    workSheet.Cells["A1:AM1"].AutoFilter = true;
                    allCells.AutoFitColumns();
                    excelPackage.Save();

                    excelPackage = null;

                    string territorios = "";
                    bool firstOnly = true;
                    for(int y = 0; y < lista_territorios.Count; y++)
                    {
                        if (firstOnly)
                        {
                            territorios = lista_territorios[y];
                            firstOnly = false;
                        }
                        else if (y + 1 == lista_territorios.Count)
                        {
                            territorios += " y " + lista_territorios[y];
                        }
                        else
                            territorios += " ," + lista_territorios[y];

                    }

                    EnvioCorreo_PdteWeb_PS_AT_TAM(p.Key, fileInfo.FullName, territorios);

                    Thread.Sleep(3000);

                    // Una vez enviado todo 
                    // borramos los archivos.

                    if (param.GetValue("Borrar_archivos") == "S")
                    {
                        fileInfo.Delete();
                    }

                    fileInfo = null;
                }


            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
                ficheroLog.AddError("CreaInformeATR: "+ ex.Message);
            }

        }

        private void EnvioCorreo_PdteWeb_PS_AT_TAM(string para, string archivo, string territorios)
        {
            FileInfo fileInfo = new FileInfo(archivo);
            StringBuilder textBody = new StringBuilder();

            try
            {
                string from = param.GetValue("Buzon_envio_email");                
                string to = para;
                string cc = "";
                string subject = "Informe ATR" + " a " + DateTime.Now.ToString("dd/MM/yyyy")
                    + " de los territorios " + territorios;

                textBody.Append(System.Environment.NewLine);
                textBody.Append(DateTime.Now.Hour < 14 ? "Buenos días:" : "Buenas tardes:");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("  Adjuntamos estado de Solicitudes ATR recibidas en el último mes. ").Append(fileInfo.Name).Append(".");
                textBody.Append(System.Environment.NewLine);
                textBody.Append("Un saludo.");

                

                if (param.GetValue("EnviarMail") == "S")
                {
                    // EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
                    EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
                    mes.SendMail(from, to, cc, subject, textBody.ToString(), archivo);
                }
                else
                {
                    office.SendMail mail_AAPP = new office.SendMail(from);
                    mail_AAPP.para.Add(para);
                    mail_AAPP.asunto = subject;
                    mail_AAPP.htmlCuerpo = textBody.ToString();
                    mail_AAPP.adjuntos.Add(archivo);
                    mail_AAPP.Save();
                }                    

                ficheroLog.Add("Correo enviado desde: " + from);
            }
            catch (Exception e)
            {
                ficheroLog.AddError("EnvioCorreo: " + e.Message);
            }
        }
    }
}
