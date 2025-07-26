using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;
using EndesaEntity;
using System.IO;
using OfficeOpenXml;
using System.Globalization;
using System.Data.Odbc;
using EndesaEntity.medida;
using EndesaBusiness.medida.Redshift;

namespace EndesaBusiness.medida
{
    public class CurvasGestores
    {
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;
        logs.Log ficheroLog;

        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.MAPIFolder outbox; // Para mover el correo
        Outlook.Items items;
        Outlook.Application myApp;

        utilidades.Param param;
        utilidades.ParamUser paramUser;

        EndesaEntity.office.ReglaCorreo regla;

        EndesaBusiness.medida.Redshift.Estados_Curvas estados_curvas;

        List<CurvaCuartoHoraria> list_cc { get; set; }
        //Dictionary<string, List<CurvaCuartoHoraria>> dic_cc { get; set; }
        List<string> lista_mails;

        bool hayError = false;
        string descripcion_error = "";

        utilidades.ZipUnZip zip;
        private enum TipoRespuestaCorreo
        {
            ConCurvas,
            SinCurvas,
            ErrorFormato
        }

        public CurvasGestores()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_CurvasGestores");
            //param = new utilidades.Param("cc_param_pruebas", servidores.MySQLDB.Esquemas.MED);
            param = new utilidades.Param("cc_param", servidores.MySQLDB.Esquemas.MED);
            //paramUser = new utilidades.ParamUser("cc_param_user_pruebas", System.Environment.UserName, servidores.MySQLDB.Esquemas.MED);
            paramUser = new utilidades.ParamUser("cc_param_user", System.Environment.UserName, servidores.MySQLDB.Esquemas.MED);
            list_cc = new List<CurvaCuartoHoraria>();
            //dic_cc = new Dictionary<string, List<CurvaCuartoHoraria>>();

            estados_curvas = new Redshift.Estados_Curvas();

            regla = new EndesaEntity.office.ReglaCorreo();

            regla.buzon = paramUser.GetValue("buzon", DateTime.Now, DateTime.Now);
            regla.asunto = param.GetValue("asunto", DateTime.Now, DateTime.Now);

            regla.moverCorreo = param.GetValue("mover_correo", DateTime.Now, DateTime.Now) == "S" ? true : false;
            regla.carpetaDestinoCorreo = param.GetValue("carpeta_buzon_destino", DateTime.Now, DateTime.Now);
            regla.de = param.GetValue("de", DateTime.Now, DateTime.Now);

            regla.carpeta = paramUser.GetValue("buzon_carpeta", DateTime.Now, DateTime.Now);
            regla.conAdjuntos = true;
            regla.guardarAdjuntos = true;

            lista_mails = new List<string>();

            DirectoryInfo dir = new DirectoryInfo(param.GetValue("ruta_temporal_archivos", DateTime.Now, DateTime.Now));
            if (!dir.Exists)
                dir.Create();

            regla.rutaSalvadoAdjuntos = param.GetValue("ruta_temporal_archivos", DateTime.Now, DateTime.Now);
            ss_pp = new EndesaBusiness.utilidades.Seguimiento_Procesos();

            zip = new utilidades.ZipUnZip();
               
        }

        public void Main()
        {
            try
            {
                ss_pp.Update_Fecha_Inicio("Medida", "Petición Curvas Gestores", "Petición Curvas Gestores");
                CompruebaBuzon();
                ss_pp.Update_Fecha_Fin("Medida", "Petición Curvas Gestores", "Petición Curvas Gestores");

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Main " + e.Message);
            }
        }

        public void CompruebaBuzon()
        {
            string email = "";
            Outlook.MailItem newEmail = null;
            List<Outlook.MailItem> moveEmail = new List<Outlook.MailItem>();
            string[] lista_de;
            bool firstOnly = true;
            string respuesta_correo = "";

            try
            {

                Console.WriteLine("Inicializando reglas de OutLook.");
                Ini_Variables_Outllook(regla);

                lista_de = regla.de.ToUpper().Split(';');

                int totalCorreos = 0;

                Console.WriteLine("Leyendo mails de " + regla.buzon);
                items.Sort("[ReceivedTime]", false);
                foreach (object eMail in items)
                {
                    firstOnly = true;
                    newEmail = eMail as Outlook.MailItem;
                    Console.Write(".");

                    if (newEmail != null)
                    {
                        if (newEmail.Subject != null)
                        {
                            Console.WriteLine(newEmail.ReceivedTime.ToString("dd/MM/yyyy") + " - " + newEmail.Subject.ToUpper());
                            // ficheroLog.Add(newEmail.ReceivedTime.ToString("dd/MM/yyyy") + " - " + newEmail.Subject.ToUpper());
                            if ((newEmail.Subject.Trim().Length >= regla.asunto.ToUpper().Length) &&
                                 regla.asunto.ToUpper() == newEmail.Subject.Trim().ToUpper().Substring(0, regla.asunto.ToUpper().Length))
                            {
                                Console.WriteLine("");
                                email = this.getSenderEmailAddress(newEmail);
                                Console.WriteLine(email);
                                if ((regla.de == null || regla.de == "") || (Contiene_de(email, regla.de)))
                                {
                                    totalCorreos++;

                                    if (regla.conAdjuntos && (newEmail.Attachments.Count > 0))
                                    {

                                        // Nos aseguramos que el directorio no tiene restos
                                        this.BorrarContenidoDirectorio(param.GetValue("ruta_temporal_archivos", DateTime.Now, DateTime.Now));

                                        #region GuardarAdjuntos
                                        if (regla.guardarAdjuntos)
                                        {
                                            for (int i = 1; i <= newEmail.Attachments.Count; i++)
                                            {
                                                // Caso que existe excel pero no tiene el nombre correcto
                                                if (newEmail.Attachments[i].FileName.ToUpper().Contains(".XLSX"))
                                                {
                                                    #region Existe_Excel_nombre_Correcto
                                                    if (newEmail.Attachments[i].FileName.ToUpper() ==
                                                       param.GetValue("nombre_archivo_adjunto", DateTime.Now, DateTime.Now).ToUpper())
                                                    {

                                                        if (firstOnly)
                                                        {
                                                            string filePath = Path.Combine(regla.rutaSalvadoAdjuntos, newEmail.Attachments[i].FileName);
                                                            newEmail.Attachments[i].SaveAsFile(filePath);
                                                            Console.WriteLine("Guardando adjunto: " + filePath);

                                                            lista_mails.Add(newEmail.EntryID.ToString());
                                                            ProcesaExcelCurvas_v2(newEmail);
                                                            firstOnly = false;

                                                            // GeneraNuevoCorreo(newEmail);
                                                            // newEmail.Move(outbox);
                                                        }
                                                    }
                                                    #endregion
                                                    else
                                                    {
                                                        respuesta_correo = "El fichero adjunto no tiene el formato correcto." + System.Environment.NewLine
                                                            + "Por favor, vuelva a enviar la petición con el formato correcto." + System.Environment.NewLine
                                                            + "Un saludo.";
                                                        lista_mails.Add(newEmail.EntryID.ToString());
                                                        this.RespondeMail(newEmail, TipoRespuestaCorreo.ErrorFormato);
                                                        // newEmail.Move(outbox);

                                                    }
                                                }

                                            }
                                            #endregion
                                        }

                                    }
                                }
                            }
                        }
                    } // 
                } // foreach (object eMail in items)


                if (totalCorreos == 0)
                {
                    //MessageBox.Show("No se han encontrado correos en el buzón: " + regla.buzon + " con el asunto " + regla.asunto + ".",
                    //"Proceso finalizado",
                    //MessageBoxButtons.OK,
                    //MessageBoxIcon.Information);
                }
                else
                {
                    //MessageBox.Show("Se han encontrado " + totalCorreos + " emails." + System.Environment.NewLine +
                    //    "Dentro de los cuales se han encontrado " + numCups + " de " + numNoCups + " CUPS en " + this.totalArchivos + " archivos.",
                    //   "Proceso de PNT´s Finalizado",
                    //MessageBoxButtons.OK,
                    //MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                ficheroLog.AddError("CompruebaBuzon " + e.Message);
            }

        }

        private void CurvasCuartoHorarias_v2(Dictionary<string, List<string>> dic,
            Outlook.MailItem mail)
        {

            int f = 0;
            int c = 0;
            bool firstOnlyOne = true;
            DateTime fechaHora = new DateTime();

            DateTime fechaDesde = new DateTime();
            DateTime fechaHasta = new DateTime();

            string[] fecha = new string[2];

            //cups.PuntosSuministro ps = new cups.PuntosSuministro();
            //cups.PuntosMedidaVigentes psv = new cups.PuntosMedidaVigentes();
            medida.Redshift.FuentesMedidaFunciones fm = new Redshift.FuentesMedidaFunciones();
            bool hayError = false;

            long total_filas_excel = 0;
            int num_hoja = 1;

            medida.Redshift.Curvas curvas = new Redshift.Curvas();

            try
            {

                //// Completamos la lista cups20 con cups13 para indices
                //ps.CompletaCups13(lc);
                //// Completamos la lista con cups15 vigentes
                //psv.CompletaCups15(lc);

                list_cc.Clear();
                curvas.dic_cc.Clear();

                foreach (KeyValuePair<string, List<string>> p in dic)
                {

                    fecha = p.Key.Split('|');
                    fechaDesde = new DateTime(Convert.ToInt32(fecha[0].Substring(0, 4)),
                       Convert.ToInt32(fecha[0].Substring(4, 2)),
                       Convert.ToInt32(fecha[0].Substring(6, 2)));

                    fechaHasta = new DateTime(Convert.ToInt32(fecha[1].Substring(0, 4)),
                       Convert.ToInt32(fecha[1].Substring(4, 2)),
                       Convert.ToInt32(fecha[1].Substring(6, 2)));



                    //ficheroLog.Add("Consultando Curva DataMart: " +
                    //        lc[w].cups20 + " - " + lc[w].cups13 + " - "
                    //        + lc[w].fd.ToString("dd/MM/yyyy") + " ~ "
                    //        + lc[w].fh.ToString("dd/MM/yyyy"));


                    hayError = curvas.GetCurvaGestor(estados_curvas, p.Value, fechaDesde, fechaHasta, estados_curvas.estados_facturados);
                    hayError = curvas.GetCurvaGestor(estados_curvas, p.Value, fechaDesde, fechaHasta, estados_curvas.estados_registrados);
                }



                int total_registros_encontrados = 0;
                foreach (KeyValuePair<string, List<CurvaCuartoHoraria>> pp in curvas.dic_cc)
                    total_registros_encontrados += pp.Value.Count;

                ficheroLog.Add("Registros encontrados: " + total_registros_encontrados);
                if (total_registros_encontrados > 0 && !hayError)
                {

                    FileInfo file = new FileInfo(regla.rutaSalvadoAdjuntos + @"\CC_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
                    ExcelPackage excelPackage = new ExcelPackage(file);
                    var workSheet = excelPackage.Workbook.Worksheets.Add("CC_1");
                    var headerCells = workSheet.Cells[1, 1, 1, 25];
                    var allCells = workSheet.Cells[1, 1, 1, 11];
                    var headerFont = headerCells.Style.Font;


                    ficheroLog.Add("Generando Excel --> " + file.Name);

                    foreach (KeyValuePair<string, List<CurvaCuartoHoraria>> pp in curvas.dic_cc)
                    {

                        if (f + (pp.Value.Count * 96) > 1048570)
                        {
                            num_hoja++;
                            workSheet = excelPackage.Workbook.Worksheets.Add("CC_" + num_hoja);
                            headerCells = workSheet.Cells[1, 1, 1, 25];
                            headerFont = headerCells.Style.Font;
                            allCells = workSheet.Cells[1, 1, f, 11];
                            allCells.AutoFitColumns();
                            firstOnlyOne = true;
                            f = 0;
                        }

                        list_cc = pp.Value.OrderBy(z => z.cups22).ThenBy(z => z.fecha).ToList();

                        for (int x = 0; x < list_cc.Count; x++)
                        {
                            total_filas_excel++;

                            Console.CursorLeft = 0;
                            Console.Write(list_cc[x].cups22 + ": " + f);
                            fechaHora = list_cc[x].fecha;




                            #region Cabecera
                            if (firstOnlyOne)
                            {
                                c = 1;
                                f++;
                                workSheet.Cells[f, c].Value = "CUPS15"; c++;
                                workSheet.Cells[f, c].Value = "CUPS22"; c++;
                                workSheet.Cells[f, c].Value = "FECHA"; c++;
                                workSheet.Cells[f, c].Value = "HORA"; c++;
                                workSheet.Cells[f, c].Value = "Energía Activa Horaria (kWh)"; c++;
                                workSheet.Cells[f, c].Value = "Energía Reactiva 1 Horaria (kVar)"; c++;
                                workSheet.Cells[f, c].Value = "Energía Reactiva 4 Horaria (kVar)"; c++;
                                workSheet.Cells[f, c].Value = "Potencia Activa"; c++;
                                workSheet.Cells[f, c].Value = "Cuarto Horaria Activa"; c++;
                                workSheet.Cells[f, c].Value = "FUENTE FINAL"; c++;
                                workSheet.Cells[f, c].Value = "ESTADO"; c++;

                                headerFont.Bold = true;
                                workSheet.View.FreezePanes(2, 1);
                                workSheet.Cells["A1:K1"].AutoFilter = true;

                                firstOnlyOne = false;

                            }
                            #endregion

                            for (int p = 1; p <= list_cc[x].numPeriodos; p++)
                            {
                                f++;
                                #region 23 Periodos        
                                if (list_cc[x].numPeriodos == 23 && p > 2)
                                {

                                    c = 1;

                                    if (p == 3)
                                        fechaHora = fechaHora.AddHours(1);

                                    workSheet.Cells[f, c].Value = list_cc[x].cups15; c++;
                                    workSheet.Cells[f, c].Value = list_cc[x].cups22; c++;
                                    workSheet.Cells[f, c].Value = fechaHora.Date;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                                    workSheet.Cells[f, c].Value = fechaHora.ToString("HH:mm:ss"); c++;
                                    workSheet.Cells[f, c].Value = list_cc[x].a[p + 1];
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                    workSheet.Cells[f, c].Value = list_cc[x].r1[p + 1];
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                    workSheet.Cells[f, c].Value = list_cc[x].r4[p + 1];
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                    workSheet.Cells[f, c].Value = list_cc[x].value[((p + 1) * 4) - 3];
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                    workSheet.Cells[f, c].Value = list_cc[x].value[((p + 1) * 4) - 3] / 4;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                                    if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                        list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                                    {
                                        workSheet.Cells[f, c].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                            Convert.ToInt32(list_cc[x].fa[p + 1])); c++;
                                    }

                                    workSheet.Cells[f, c].Value = list_cc[x].estado; c++;


                                    fechaHora = fechaHora.AddMinutes(15);
                                    f++;

                                    workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                    workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                    workSheet.Cells[f, 3].Value = fechaHora.Date;
                                    workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                    workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4) - 2];
                                    workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 9].Value = list_cc[x].value[((p + 1) * 4) - 2] / 4;
                                    workSheet.Cells[f, 9].Style.Numberformat.Format = "#,##0";

                                    if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                        list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                                    {
                                        workSheet.Cells[f, 10].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                            Convert.ToInt32(list_cc[x].fa[p + 1]));
                                    }

                                    workSheet.Cells[f, 11].Value = list_cc[x].estado;

                                    fechaHora = fechaHora.AddMinutes(15);
                                    f++;

                                    workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                    workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                    workSheet.Cells[f, 3].Value = fechaHora.Date;
                                    workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                    workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4) - 1];
                                    workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 9].Value = list_cc[x].value[((p + 1) * 4) - 1] / 4;
                                    workSheet.Cells[f, 9].Style.Numberformat.Format = "#,##0";

                                    if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                        list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                                    {
                                        workSheet.Cells[f, 10].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                            Convert.ToInt32(list_cc[x].fa[p + 1]));
                                    }

                                    workSheet.Cells[f, 11].Value = list_cc[x].estado;

                                    fechaHora = fechaHora.AddMinutes(15);
                                    f++;

                                    workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                    workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                    workSheet.Cells[f, 3].Value = fechaHora.Date;
                                    workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                    workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4)];
                                    workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 9].Value = list_cc[x].value[((p + 1) * 4)] / 4;
                                    workSheet.Cells[f, 9].Style.Numberformat.Format = "#,##0";
                                    fechaHora = fechaHora.AddMinutes(15);

                                    if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                        list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                                    {
                                        workSheet.Cells[f, 10].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                            Convert.ToInt32(list_cc[x].fa[p + 1]));
                                    }

                                    workSheet.Cells[f, 11].Value = list_cc[x].estado;
                                }
                                #endregion
                                #region 25 Periodos
                                else if (list_cc[x].numPeriodos == 25 && p > 2)
                                {
                                    if (p == 4)
                                        fechaHora = fechaHora.AddHours(-1);

                                    ficheroLog.Add("p: " + p + " x: " + x);

                                    workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                    workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                    workSheet.Cells[f, 3].Value = fechaHora.Date;
                                    workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                    workSheet.Cells[f, 5].Value = list_cc[x].a[p];
                                    workSheet.Cells[f, 5].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 6].Value = list_cc[x].r1[p];
                                    workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 7].Value = list_cc[x].r4[p];
                                    workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 3];
                                    workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 9].Value = list_cc[x].value[(p * 4) - 3] / 4;
                                    workSheet.Cells[f, 9].Style.Numberformat.Format = "#,##0";

                                    //if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                    //    list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                                    if (list_cc[x].fc[p] != null && list_cc[x].fc[p] != "" &&
                                        list_cc[x].fa[p] != null && list_cc[x].fa[p] != "")
                                    {

                                        workSheet.Cells[f, 10].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                            Convert.ToInt32(list_cc[x].fa[p]));
                                    }

                                    workSheet.Cells[f, 11].Value = list_cc[x].estado;

                                    fechaHora = fechaHora.AddMinutes(15);
                                    f++;

                                    workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                    workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                    workSheet.Cells[f, 3].Value = fechaHora.Date;
                                    workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                    workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 2];
                                    workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 9].Value = list_cc[x].value[(p * 4) - 2] / 4;
                                    workSheet.Cells[f, 9].Style.Numberformat.Format = "#,##0";

                                    if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                        (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                    {

                                        workSheet.Cells[f, 10].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                            Convert.ToInt32(list_cc[x].fa[p]));
                                    }

                                    workSheet.Cells[f, 11].Value = list_cc[x].estado;

                                    fechaHora = fechaHora.AddMinutes(15);
                                    f++;

                                    workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                    workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                    workSheet.Cells[f, 3].Value = fechaHora.Date;
                                    workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                    workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 1];
                                    workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 9].Value = list_cc[x].value[(p * 4) - 1] / 4;
                                    workSheet.Cells[f, 9].Style.Numberformat.Format = "#,##0";

                                    if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                        (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                    {

                                        workSheet.Cells[f, 10].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                            Convert.ToInt32(list_cc[x].fa[p]));
                                    }

                                    workSheet.Cells[f, 11].Value = list_cc[x].estado;

                                    fechaHora = fechaHora.AddMinutes(15);
                                    f++;

                                    workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                    workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                    workSheet.Cells[f, 3].Value = fechaHora.Date;
                                    workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                    workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4)];
                                    workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 9].Value = list_cc[x].value[(p * 4)] / 4;
                                    workSheet.Cells[f, 9].Style.Numberformat.Format = "#,##0";
                                    fechaHora = fechaHora.AddMinutes(15);

                                    if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                        (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                    {
                                        workSheet.Cells[f, 10].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]), Convert.ToInt32(list_cc[x].fa[p]));
                                    }

                                    workSheet.Cells[f, 11].Value = list_cc[x].estado;

                                }
                                #endregion
                                #region 24 Periodos
                                else
                                {
                                    workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                    workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                    workSheet.Cells[f, 3].Value = fechaHora.Date;
                                    workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                    workSheet.Cells[f, 5].Value = list_cc[x].a[p];
                                    workSheet.Cells[f, 5].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 6].Value = list_cc[x].r1[p];
                                    workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 7].Value = list_cc[x].r4[p];
                                    workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 3];
                                    workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 9].Value = list_cc[x].value[(p * 4) - 3] / 4;
                                    workSheet.Cells[f, 9].Style.Numberformat.Format = "#,##0";

                                    if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                        (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                    {
                                        workSheet.Cells[f, 10].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                            Convert.ToInt32(list_cc[x].fa[p]));
                                    }

                                    workSheet.Cells[f, 11].Value = list_cc[x].estado;

                                    fechaHora = fechaHora.AddMinutes(15);
                                    f++;
                                    workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                    workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                    workSheet.Cells[f, 3].Value = fechaHora.Date;
                                    workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                    workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 2];
                                    workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 9].Value = list_cc[x].value[(p * 4) - 2] / 4;
                                    workSheet.Cells[f, 9].Style.Numberformat.Format = "#,##0";

                                    if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                        (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                    {
                                        workSheet.Cells[f, 10].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                            Convert.ToInt32(list_cc[x].fa[p]));
                                    }

                                    workSheet.Cells[f, 11].Value = list_cc[x].estado;

                                    fechaHora = fechaHora.AddMinutes(15);
                                    f++;
                                    workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                    workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                    workSheet.Cells[f, 3].Value = fechaHora.Date;
                                    workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                    workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 1];
                                    workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 9].Value = list_cc[x].value[(p * 4) - 1] / 4;
                                    workSheet.Cells[f, 9].Style.Numberformat.Format = "#,##0";


                                    if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                        (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                    {
                                        workSheet.Cells[f, 10].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                            Convert.ToInt32(list_cc[x].fa[p]));
                                    }

                                    workSheet.Cells[f, 11].Value = list_cc[x].estado;

                                    fechaHora = fechaHora.AddMinutes(15);
                                    f++;
                                    workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                    workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                    workSheet.Cells[f, 3].Value = fechaHora.Date;
                                    workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                    workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                    workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4)];
                                    workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                    workSheet.Cells[f, 9].Value = list_cc[x].value[(p * 4)] / 4;
                                    workSheet.Cells[f, 9].Style.Numberformat.Format = "#,##0";
                                    fechaHora = fechaHora.AddMinutes(15);

                                    if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                        (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                    {
                                        workSheet.Cells[f, 10].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                            Convert.ToInt32(list_cc[x].fa[p]));
                                    }

                                    workSheet.Cells[f, 11].Value = list_cc[x].estado;
                                }

                                #endregion
                            }
                        }
                    }


                    allCells = workSheet.Cells[1, 1, f, 11];
                    allCells.AutoFitColumns();
                    excelPackage.Save();
                    excelPackage = null;

                    Console.WriteLine("Comprimiendo archivo ...");
                    file = new FileInfo(file.FullName);

                    double size = file.Length / 1024; // KB
                    if (size > 10000)
                        zip.ComprimirArchivo_Split(file.FullName, file.FullName.Replace(".xlsx", ".zip"), 10);
                    else
                        zip.ComprimirArchivo(file.FullName, file.FullName.Replace(".xlsx", ".zip"));

                    file.Delete();

                    string[] listaArchivos = Directory.GetFiles(regla.rutaSalvadoAdjuntos,
                       param.GetValue("prefijo_archivo_curvas", DateTime.Now, DateTime.Now));
                    if (listaArchivos.Count() == 1)
                        this.RespondeMail(mail, TipoRespuestaCorreo.ConCurvas);
                    else
                    {
                        for (int i = 0; i < listaArchivos.Count(); i++)
                            this.RespondeMailMultiArchivo(mail,
                                TipoRespuestaCorreo.ConCurvas, listaArchivos[i], i + 1, listaArchivos.Count());

                        mail.Move(outbox);
                    }



                }
                else if (!hayError)// Caso en el que no se han encontrado datos
                {
                    // this.RespondeCorreo(mail, "No existen curvas facturadas para los cups y periodos solicitados.");
                    this.RespondeMail(mail, TipoRespuestaCorreo.SinCurvas);

                }

            }
            catch (Exception e)
            {
                ficheroLog.AddError("CurvasRedShift.CurvasCuartoHorarias_v2: "
                    + e.Message + " --> " + this.getSenderEmailAddress(mail)
                    + " Fila: " + f + "FechaHora: " + fechaHora);
            }
        }

        

        private bool Contiene_de(string valor, string texto)
        {
            // True si la lista texto con separador ; aparece en el texto valor
            string[] lista;
            lista = texto.Split(';');

            for (int i = 0; i < lista.Count(); i++)
            {
                if (valor.ToUpper().Contains(lista[i].ToUpper()))
                {
                    return true;
                }
            }
            return false;
        }
        private void ProcesaExcelCurvas_v2(Outlook.MailItem mail)
        {
            string[] listaArchivos;
            FileInfo file;
            string respuesta_correo;
            Dictionary<string, List<string>> dic;
            try
            {
                listaArchivos = Directory.GetFiles(regla.rutaSalvadoAdjuntos,
                    param.GetValue("prefijo_archivo_solicitud_curvas", DateTime.Now, DateTime.Now));

                for (int i = 0; i < listaArchivos.Count(); i++)
                {
                    file = new FileInfo(listaArchivos[i]);
                    dic = LeeExcel_Peticion(file.FullName, mail);

                    if (!hayError)
                    {
                        this.CurvasCuartoHorarias_v2(dic, mail);
                    }
                    else
                    {
                        // Caso que el excel no cumple con alguna validación
                        respuesta_correo = "El Excel adjunto no tiene el formato correcto." + System.Environment.NewLine
                            + "Por favor, vuelva a enviar la petición con el formato correcto." + System.Environment.NewLine
                            + "Un saludo.";
                        // this.RespondeCorreo(mail,respuesta_correo);
                        this.RespondeMail(mail, TipoRespuestaCorreo.ErrorFormato);
                    }

                    file.Delete();
                }

            }
            catch (Exception e)
            {
                ficheroLog.AddError("CurvasNetezza.ProcesaExcelCurvas " + e.Message);
            }
        }

        private void Ini_Variables_Outllook(EndesaEntity.office.ReglaCorreo regla)
        {
            myApp = new Outlook.Application();
            outlookNameSpace = myApp.GetNamespace("MAPI");

            try
            {
                if (regla.buzon != null)
                    inbox = outlookNameSpace.Folders[regla.buzon].Folders[regla.carpeta];
                else
                    inbox = outlookNameSpace.GetDefaultFolder(
                        Microsoft.Office.Interop.Outlook.
                        OlDefaultFolders.olFolderInbox);

                if (regla.moverCorreo)
                    outbox = outlookNameSpace.Folders[regla.buzon].Folders[regla.carpeta].Folders[regla.carpetaDestinoCorreo];

                items = inbox.Items;

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Ini_Variables_Outllook: " + e.Message);
                ficheroLog.AddError("regla.buzon: " + regla.buzon);
                ficheroLog.AddError("regla.carpeta: " + regla.carpeta);
                ficheroLog.AddError("regla.carpetaDestinoCorreo: " + regla.carpetaDestinoCorreo);
            }

        }

        private string getSenderEmailAddress(Outlook.MailItem mail)
        {
            Outlook.AddressEntry sender = mail.Sender;
            string SenderEmailAddress = "";

            if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry ||
                sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
            {
                Outlook.ExchangeUser exchUser = sender.GetExchangeUser();
                if (exchUser != null)
                {
                    SenderEmailAddress = exchUser.PrimarySmtpAddress;
                }
            }
            else
            {
                SenderEmailAddress = mail.SenderEmailAddress;
            }

            return SenderEmailAddress;
        }

        private void BorrarContenidoDirectorio(string directorio)
        {
            string[] listaArchivos;
            FileInfo file;

            listaArchivos = Directory.GetFiles(regla.rutaSalvadoAdjuntos);
            for (int i = 0; i < listaArchivos.Count(); i++)
            {
                file = new FileInfo(listaArchivos[i]);
                file.Delete();
            }
        }        

        private string GetCCEmailAddress(Outlook.MailItem mail)
        {
            string email;
            string lista_cc = "";
            bool firstOnly = true;

            Outlook.ExchangeUser exUser;
            List<string> ccEmailAddressList = new List<string>();
            foreach (Outlook.Recipient recip in mail.Recipients)
            {
                if ((Outlook.OlMailRecipientType)recip.Type == Outlook.OlMailRecipientType.olCC)
                {
                    email = recip.Address;
                    if (!email.Contains("@"))
                    {
                        exUser = recip.AddressEntry.GetExchangeUser();
                        email = exUser.PrimarySmtpAddress;
                    }
                    ccEmailAddressList.Add(email);

                }
            }

            for (int i = 0; i < ccEmailAddressList.Count(); i++)
            {
                if (firstOnly)
                {
                    lista_cc = ccEmailAddressList[i];
                    firstOnly = false;
                }
                else
                {
                    lista_cc += ";" + ccEmailAddressList[i];
                }
            }
            return lista_cc;
        }

        private string CuerpoMailCustom(string value)
        {
            string saludo = "";
            string texto = "";

            if (DateTime.Now.Hour > 13)
                saludo = "Buenas tardes:";
            else
                saludo = "Buenos días:";

            texto = param.GetValue(value, DateTime.Now, DateTime.Now).Replace("Buenos días", saludo);
            // texto = texto.Replace("endesa.png", param.GetValue("mail_image_path", DateTime.Now, DateTime.Now));

            return texto;
        }

        private Dictionary<string, List<string>> LeeExcel_Peticion(string archivo, Outlook.MailItem mail)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage;
            FileStream fs;
            int c = 1;
            int f = 1;
            int id = 0;
            bool firstOnly = true;
            string cabecera = "";
            string fecha = "";

            curvasDataMart.SolicitudFunciones sf = new curvasDataMart.SolicitudFunciones();
            curvasDataMart.SolictudDetalleFunciones sfd = new curvasDataMart.SolictudDetalleFunciones();

            Dictionary<string, List<string>> d = new Dictionary<string, List<string>>();

            try
            {

                sf.mail = getSenderEmailAddress(mail);
                sf.fechahora_mail = mail.ReceivedTime;

                fs = new FileStream(archivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();
                f = 1; // Porque la primera fila es la cabecera
                for (int i = 0; i < 157; i++)
                {
                    c = 1;

                    if (firstOnly)
                    {
                        cabecera = workSheet.Cells[1, 1].Value.ToString()
                            + workSheet.Cells[1, 2].Value.ToString()
                            + workSheet.Cells[1, 3].Value.ToString();

                        if (!EstructuraCorrecta(cabecera))
                        {
                            hayError = true;
                            descripcion_error = "La estructura del archivo excel no es la correcta.";
                            sf.desc_error = "La estructura del archivo excel no es la correcta.";
                            sf.Save();
                            break;
                        }
                        else
                        {

                            sf.Save();
                        }
                        firstOnly = false;
                    }

                    f++;

                    if (workSheet.Cells[f, 1].Value == null
                        || workSheet.Cells[f, 2].Value == null
                        || workSheet.Cells[f, 3].Value == null)
                    {
                        break;
                    }
                    else
                    {

                        EndesaEntity.cups.PuntoSuministro cups = new EndesaEntity.cups.PuntoSuministro();
                        cups.id = id;
                        cups.cups20 = workSheet.Cells[f, c].Value.ToString().Trim().Substring(0, 20); c++;
                        cups.fd = Convert.ToDateTime(workSheet.Cells[f, c].Value.ToString().Trim()); c++;
                        cups.fh = Convert.ToDateTime(workSheet.Cells[f, c].Value.ToString().Trim());
                        if (cups.fd <= cups.fh && (cups.cups20.Length >= 20 && cups.cups20.Length <= 22))
                        {
                            id++;

                            fecha = cups.fd.ToString("yyyyMMdd") + "|" + cups.fh.ToString("yyyyMMdd");

                            List<string> o;
                            if (!d.TryGetValue(fecha, out o))
                            {
                                o = new List<string>();
                                o.Add(cups.cups20);
                                d.Add(fecha, o);
                            }
                            else
                                o.Add(cups.cups20);

                            sfd.id = sf.id;
                            sfd.cups20 = cups.cups20;
                            sfd.fd = cups.fd;
                            sfd.fh = cups.fh;
                            sfd.Save();
                        }

                    }

                }

                fs.Close();
                fs = null;
                excelPackage = null;


                return d;
            }
            catch (Exception e)
            {
                ficheroLog.AddError("CargaExcel: " + e.Message);
                hayError = true;
                return null;
            }




        }

        private bool EstructuraCorrecta(string cabecera)
        {
            // return cabecera == "CUPS20FECHA DESDEFECHA HASTA";
            return true;
        }

        private void RespondeMail(Outlook.MailItem mail, TipoRespuestaCorreo respuesta)
        {

            string[] listaArchivos;
            FileInfo file;
            Outlook.MailItem theMail;
            string enCopia = "";
            string para = "";
            string cc = "";
            bool firstOnly = true;
            try
            {

                theMail = mail.Reply();
                //para = theMail.To;
                para = getSenderEmailAddress(mail);
                theMail.To = getSenderEmailAddress(mail);



                // Comprobamos que tengamos adjuntos para crear el correo
                listaArchivos = Directory.GetFiles(regla.rutaSalvadoAdjuntos,
                    param.GetValue("prefijo_archivo_curvas", DateTime.Now, DateTime.Now));

                switch (respuesta)
                {
                    case TipoRespuestaCorreo.ConCurvas:
                        theMail.HTMLBody = this.CuerpoMailCustom("html_body") + theMail.HTMLBody;
                        break;
                    case TipoRespuestaCorreo.SinCurvas:
                        theMail.HTMLBody = this.CuerpoMailCustom("html_body_sin_datos") + theMail.HTMLBody;
                        break;
                    default:
                        theMail.HTMLBody = this.CuerpoMailCustom("html_body_error_formato") + theMail.HTMLBody;
                        break;
                }

                if (mail.CC != null)
                {
                    theMail.CC = GetCCEmailAddress(mail);
                    enCopia = GetCCEmailAddress(mail);
                }



                for (int i = 0; i < listaArchivos.Count(); i++)
                {
                    theMail.Attachments.Add(listaArchivos[i]);
                }

                theMail.Attachments.Add(System.Environment.CurrentDirectory + param.GetValue("mail_image_path", DateTime.Now, DateTime.Now));
                if (param.GetValue("adjuntar_correo_ejemplo", DateTime.Now, DateTime.Now) == "S")
                {
                    theMail.Attachments.Add(System.Environment.CurrentDirectory + param.GetValue("correo_ejemplo", DateTime.Now, DateTime.Now));
                }

                ficheroLog.Add("Enviando correo Para: " + para + " CC:" + enCopia);

                if (paramUser.GetValue("enviar_mail", DateTime.Now, DateTime.Now) == "N")
                    theMail.Save();
                else
                    theMail.Send();

                Thread.Sleep(1000);


                for (int i = 0; i < listaArchivos.Count(); i++)
                {
                    file = new FileInfo(listaArchivos[i]);
                    file.Delete();
                }

                mail.Move(outbox);
            }
            catch (Exception e)
            {
                ficheroLog.AddError("CurvasNetezza.GeneraNuevoCorreo: " + e.Message + " En copia: " + enCopia + " Para: ");
            }
        }

        private void RespondeMailMultiArchivo(Outlook.MailItem mail, TipoRespuestaCorreo respuesta,
            string archivo, int num_mail, int total_mails)
        {


            Outlook.MailItem theMail;
            FileInfo file;
            string enCopia = "";
            string para = "";
            string cc = "";
            bool firstOnly = true;
            try
            {

                theMail = mail.Reply();
                //para = theMail.To;
                para = getSenderEmailAddress(mail);
                theMail.To = getSenderEmailAddress(mail);
                theMail.Subject = theMail.Subject
                    + " "
                    + num_mail + " de " + total_mails;



                switch (respuesta)
                {
                    case TipoRespuestaCorreo.ConCurvas:
                        theMail.HTMLBody = this.CuerpoMailCustom("html_body") + theMail.HTMLBody;
                        break;
                    case TipoRespuestaCorreo.SinCurvas:
                        theMail.HTMLBody = this.CuerpoMailCustom("html_body_sin_datos") + theMail.HTMLBody;
                        break;
                    default:
                        theMail.HTMLBody = this.CuerpoMailCustom("html_body_error_formato") + theMail.HTMLBody;
                        break;
                }

                if (mail.CC != null)
                {
                    theMail.CC = GetCCEmailAddress(mail);
                    enCopia = GetCCEmailAddress(mail);
                }





                theMail.Attachments.Add(archivo);


                theMail.Attachments.Add(System.Environment.CurrentDirectory + param.GetValue("mail_image_path", DateTime.Now, DateTime.Now));
                if (param.GetValue("adjuntar_correo_ejemplo", DateTime.Now, DateTime.Now) == "S")
                {
                    theMail.Attachments.Add(System.Environment.CurrentDirectory + param.GetValue("correo_ejemplo", DateTime.Now, DateTime.Now));
                }

                ficheroLog.Add("Enviando correo Para: " + para + " CC:" + enCopia);

                if (paramUser.GetValue("enviar_mail", DateTime.Now, DateTime.Now) == "N")
                    theMail.Save();
                else
                    theMail.Send();

                Thread.Sleep(1000);




                file = new FileInfo(archivo);
                file.Delete();


            }
            catch (Exception e)
            {
                ficheroLog.AddError("CurvasNetezza.GeneraNuevoCorreo: " + e.Message + " En copia: " + enCopia + " Para: ");
            }
        }
    }
}

