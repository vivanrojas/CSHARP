using EndesaBusiness.servidores;
using Microsoft.Exchange.WebServices.Data;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class CurvasGestoresBuzon
    {

        List<string> lista_mails = new List<string>();
        private enum TipoRespuestaCorreo
        {
            ConCurvas,
            SinCurvas,
            ErrorFormato
        }


        ExchangeService service;
        Mailbox mb;
        FolderId fid;
        
        Folder inbox;
        FolderId outbox;

        utilidades.Param param;
        utilidades.ParamUser paramUser;

        List<EndesaEntity.medida.CurvaCuartoHorariaInformes> list_cc { get; set; }
        List<EndesaEntity.medida.PuntoSuministro> lc = new List<EndesaEntity.medida.PuntoSuministro>();

        EndesaEntity.ReglaCorreo regla;

        logs.Log ficheroLog;
        public CurvasGestoresBuzon()
        {

            utilidades.Credenciales credenciales = new utilidades.Credenciales("ES02255021D");

            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_CurvasGestores");
            param = new utilidades.Param("cc_param", servidores.MySQLDB.Esquemas.MED);
            paramUser = new utilidades.ParamUser("cc_param_user", "ES44977034M", servidores.MySQLDB.Esquemas.MED);


            list_cc = new List<EndesaEntity.medida.CurvaCuartoHorariaInformes>();


            service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials(credenciales.server_user, credenciales.server_password);
            service.AutodiscoverUrl(credenciales.server_user, RedirectionUrlValidationCallback);
            mb = new Mailbox(paramUser.GetValue("buzon", DateTime.Now, DateTime.Now));
            fid = new FolderId(WellKnownFolderName.Inbox, mb);
            inbox = Folder.Bind(service, fid);

            outbox = GetFolderID(param.GetValue("carpeta_buzon_destino", DateTime.Now, DateTime.Now));


            regla = new EndesaEntity.ReglaCorreo();

            regla.buzon = paramUser.GetValue("buzon", DateTime.Now, DateTime.Now);
            regla.asunto = param.GetValue("asunto", DateTime.Now, DateTime.Now);

            regla.moverCorreo = param.GetValue("mover_correo", DateTime.Now, DateTime.Now) == "S" ? true : false;
            regla.carpetaDestinoCorreo = param.GetValue("carpeta_buzon_destino", DateTime.Now, DateTime.Now);
            regla.de = param.GetValue("de", DateTime.Now, DateTime.Now);

            regla.carpeta = paramUser.GetValue("buzon_carpeta", DateTime.Now, DateTime.Now);
            regla.conAdjuntos = true;
            regla.guardarAdjuntos = true;

            DirectoryInfo dir = new DirectoryInfo(param.GetValue("ruta_temporal_archivos", DateTime.Now, DateTime.Now));
            if (!dir.Exists)
                dir.Create();

            regla.rutaSalvadoAdjuntos = param.GetValue("ruta_temporal_archivos", DateTime.Now, DateTime.Now);

        }

        public void RecorreInbox()
        {
            EmailMessage newEmail = null;
            ItemView view = new ItemView(1000, 0, OffsetBasePoint.Beginning);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(ItemSchema.DateTimeReceived);
            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);

            FindItemsResults<Item> items = inbox.FindItems(view);

            bool firstOnly = true;
            string respuesta_correo = "";

            foreach (EmailMessage eMail in items)
            {
                newEmail = EmailMessage.Bind(service, eMail.Id, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments, ItemSchema.HasAttachments));
                newEmail.Load();

                Console.Clear();
                Console.CursorLeft = 0;
                Console.Write(newEmail.DateTimeReceived.ToString("dd/MM/yyyy HH:mm:ss") + " " + newEmail.Sender.Address + " - " + newEmail.Subject);


                if ((regla.asunto == "" || (newEmail.Subject.Trim().Length >= regla.asunto.Length) &&
                                regla.ignorar_asunto || (newEmail.Subject.Trim().ToUpper().Substring(0, regla.asunto.Length) == regla.asunto.ToUpper())))
                {
                    if ((regla.de == null || regla.de == "") || (newEmail.Sender.Address.ToUpper() == regla.de.ToUpper()))
                    {

                        EndesaEntity.Correo c = new EndesaEntity.Correo();
                        c.sender = newEmail.Sender.Address;
                        c.subject = newEmail.Subject;
                        c.receivedTime = newEmail.DateTimeReceived;

                        foreach (Attachment attachment in newEmail.Attachments)
                        {
                            if (attachment.Name.ToUpper().Contains(".XLSX"))
                            {
                                string filePath = Path.Combine(regla.rutaSalvadoAdjuntos, attachment.Name);
                                FileAttachment fileAttachment = attachment as FileAttachment;
                                fileAttachment.Load(filePath);

                                string[] listaArchivos = Directory.GetFiles(regla.rutaSalvadoAdjuntos, "*.XLSX");
                                for (int i = 0; i < listaArchivos.Count(); i++)
                                {
                                    FileInfo archivo = new FileInfo(listaArchivos[i]);
                                    if (archivo.Name.ToUpper().Contains(".XLSX"))
                                    {
                                        #region Existe_Excel_nombre_Correcto
                                        if (archivo.Name.ToUpper() ==
                                           param.GetValue("nombre_archivo_adjunto", DateTime.Now, DateTime.Now).ToUpper())
                                        {

                                            if (firstOnly)
                                            {
                                                                                                
                                                Console.WriteLine("Guardando adjunto: " + filePath);                                                
                                                ProcesaExcelCurvas(newEmail);
                                                firstOnly = false;
                                                                                                
                                            }
                                        }
                                        #endregion
                                        else
                                        {
                                            respuesta_correo = "El fichero adjunto no tiene el formato correcto." + System.Environment.NewLine
                                                + "Por favor, vuelva a enviar la petición con el formato correcto." + System.Environment.NewLine
                                                + "Un saludo.";
                                            
                                            RespondeMail(newEmail, TipoRespuestaCorreo.ErrorFormato);
                                            

                                        }
                                    }

                                }
                            }


                        }

                    }
                }
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        private void ProcesaExcelCurvas(EmailMessage mail)
        {
            string[] listaArchivos;
            FileInfo file;
            string respuesta_correo;
            try
            {
                listaArchivos = Directory.GetFiles(regla.rutaSalvadoAdjuntos,
                    param.GetValue("prefijo_archivo_solicitud_curvas", DateTime.Now, DateTime.Now));

                for (int i = 0; i < listaArchivos.Count(); i++)
                {
                    file = new FileInfo(listaArchivos[i]);
                    medida.ExcelCUPS ex = new medida.ExcelCUPS(listaArchivos[i], mail);
                    if (!ex.hayError)
                    {
                        lc.Clear();
                        lc = ex.lista_cups;
                        CurvasCuartoHorarias(lc, mail);
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

        private void CurvasCuartoHorarias(List<EndesaEntity.medida.PuntoSuministro> lc, EmailMessage mail)
        {

            int f = 0;

            bool firstOnlyOne = true;

            DateTime fechaHora = new DateTime();

            cups.PuntosSuministro ps = new cups.PuntosSuministro();
            PuntosMedidaPrincipalesVigentes psv = new PuntosMedidaPrincipalesVigentes(lc);
            medida.FuentesMedida fm = new medida.FuentesMedida();
            bool hayError = false;
            utilidades.ZipUnZip zip = new utilidades.ZipUnZip();

            try
            {

                // Completamos la lista cups20 con cups13 para indices
                ps.CompletaCups13(lc);
                // Completamos la lista con cups15 vigentes
                psv.CompletaCups15(lc);

                list_cc.Clear();

                for (int w = 0; w < lc.Count; w++)
                {
                    if (lc[w].cups13 != null)
                    {
                        ficheroLog.Add("Consultando Curva DataMart: " +
                            lc[w].cups20 + " - " + lc[w].cups13 + " - "
                            + lc[w].fd.ToString("dd/MM/yyyy") + " ~ "
                            + lc[w].fh.ToString("dd/MM/yyyy"));

                        if (!hayError)
                        {

                            // Consultas directas a Datamark
                            hayError = GetCurva(lc[w].cups13, lc[w].fd, lc[w].fh, "F");
                            for (int i = 0; i < lc[w].cups15.Count; i++)
                            {
                                hayError = GetCurva(lc[w].cups15[i], lc[w].fd, lc[w].fh, "R");
                            }

                        }

                    }

                }

                ficheroLog.Add("Registros encontrados: " + list_cc.Count);
                if (list_cc.Count() > 0 && !hayError)
                {

                    FileInfo file = new FileInfo(regla.rutaSalvadoAdjuntos + @"\CC_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
                    ExcelPackage excelPackage = new ExcelPackage(file);
                    var workSheet = excelPackage.Workbook.Worksheets.Add("CC_1");
                    var headerCells = workSheet.Cells[1, 1, 1, 25];
                    var headerFont = headerCells.Style.Font;

                    ficheroLog.Add("Generando Excel --> " + file.Name);

                    list_cc = list_cc.OrderBy(z => z.cups15).ThenBy(z => z.fecha).ToList();

                    for (int x = 0; x < list_cc.Count; x++)
                    {
                        fechaHora = list_cc[x].fecha;

                        #region Cabecera
                        if (firstOnlyOne)
                        {
                            f++;
                            workSheet.Cells[f, 1].Value = "CUPS15";
                            workSheet.Cells[f, 2].Value = "CUPS22";
                            workSheet.Cells[f, 3].Value = "FECHA";
                            workSheet.Cells[f, 4].Value = "HORA";
                            workSheet.Cells[f, 5].Value = "Energía Activa Horaria (kWh)";
                            workSheet.Cells[f, 6].Value = "Energía Reactiva Horaria (kVar)";
                            workSheet.Cells[f, 7].Value = "Potencia Activa";
                            workSheet.Cells[f, 8].Value = "Cuarto Horaria Activa";
                            workSheet.Cells[f, 9].Value = "FUENTE FINAL";
                            workSheet.Cells[f, 10].Value = "ESTADO";

                            firstOnlyOne = false;

                        }
                        #endregion

                        for (int p = 1; p <= list_cc[x].numPeriodos; p++)
                        {
                            f++;
                            #region 23 Periodos        
                            if (list_cc[x].numPeriodos == 23 && p > 2)
                            {
                                if (p == 3)
                                    fechaHora = fechaHora.AddHours(1);

                                workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                workSheet.Cells[f, 3].Value = fechaHora.Date;
                                workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                workSheet.Cells[f, 5].Value = list_cc[x].a[p + 1];
                                workSheet.Cells[f, 5].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 6].Value = list_cc[x].r[p + 1];
                                workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 7].Value = list_cc[x].value[((p + 1) * 4) - 3];
                                workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4) - 3] / 4;
                                workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                                if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                    list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                                {
                                    workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                        Convert.ToInt32(list_cc[x].fa[p + 1]));
                                }

                                workSheet.Cells[f, 10].Value = list_cc[x].estado;


                                fechaHora = fechaHora.AddMinutes(15);
                                f++;
                                workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                workSheet.Cells[f, 3].Value = fechaHora.Date;
                                workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                workSheet.Cells[f, 7].Value = list_cc[x].value[((p + 1) * 4) - 2];
                                workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4) - 2] / 4;
                                workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                                if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                    list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                                {
                                    workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                        Convert.ToInt32(list_cc[x].fa[p + 1]));
                                }

                                workSheet.Cells[f, 10].Value = list_cc[x].estado;

                                fechaHora = fechaHora.AddMinutes(15);
                                f++;
                                workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                workSheet.Cells[f, 3].Value = fechaHora.Date;
                                workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                workSheet.Cells[f, 7].Value = list_cc[x].value[((p + 1) * 4) - 1];
                                workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4) - 1] / 4;
                                workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                                if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                    list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                                {
                                    workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                        Convert.ToInt32(list_cc[x].fa[p + 1]));
                                }

                                workSheet.Cells[f, 10].Value = list_cc[x].estado;

                                fechaHora = fechaHora.AddMinutes(15);
                                f++;
                                workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                workSheet.Cells[f, 3].Value = fechaHora.Date;
                                workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                workSheet.Cells[f, 7].Value = list_cc[x].value[((p + 1) * 4)];
                                workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4)] / 4;
                                workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                fechaHora = fechaHora.AddMinutes(15);

                                if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                    list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                                {
                                    workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                                        Convert.ToInt32(list_cc[x].fa[p + 1]));
                                }

                                workSheet.Cells[f, 10].Value = list_cc[x].estado;
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
                                workSheet.Cells[f, 6].Value = list_cc[x].r[p];
                                workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 3];
                                workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 3] / 4;
                                workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                                //if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                                //    list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                                if (list_cc[x].fc[p] != null && list_cc[x].fc[p] != "" &&
                                    list_cc[x].fa[p] != null && list_cc[x].fa[p] != "")
                                {

                                    workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                        Convert.ToInt32(list_cc[x].fa[p]));
                                }

                                workSheet.Cells[f, 10].Value = list_cc[x].estado;

                                fechaHora = fechaHora.AddMinutes(15);
                                f++;
                                workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                workSheet.Cells[f, 3].Value = fechaHora.Date;
                                workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 2];
                                workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 2] / 4;
                                workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                                if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                    (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                {

                                    workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                        Convert.ToInt32(list_cc[x].fa[p]));
                                }

                                workSheet.Cells[f, 10].Value = list_cc[x].estado;

                                fechaHora = fechaHora.AddMinutes(15);
                                f++;
                                workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                workSheet.Cells[f, 3].Value = fechaHora.Date;
                                workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 1];
                                workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 1] / 4;
                                workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                                if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                    (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                {

                                    workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                        Convert.ToInt32(list_cc[x].fa[p]));
                                }

                                workSheet.Cells[f, 10].Value = list_cc[x].estado;

                                fechaHora = fechaHora.AddMinutes(15);
                                f++;
                                workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                workSheet.Cells[f, 3].Value = fechaHora.Date;
                                workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4)];
                                workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4)] / 4;
                                workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                fechaHora = fechaHora.AddMinutes(15);

                                if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                    (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                {
                                    workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]), Convert.ToInt32(list_cc[x].fa[p]));
                                }

                                workSheet.Cells[f, 10].Value = list_cc[x].estado;

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
                                workSheet.Cells[f, 6].Value = list_cc[x].r[p];
                                workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 3];
                                workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 3] / 4;
                                workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                                if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                    (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                {
                                    workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                        Convert.ToInt32(list_cc[x].fa[p]));
                                }

                                workSheet.Cells[f, 10].Value = list_cc[x].estado;

                                fechaHora = fechaHora.AddMinutes(15);
                                f++;
                                workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                workSheet.Cells[f, 3].Value = fechaHora.Date;
                                workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 2];
                                workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 2] / 4;
                                workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                                if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                    (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                {
                                    workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                        Convert.ToInt32(list_cc[x].fa[p]));
                                }

                                workSheet.Cells[f, 10].Value = list_cc[x].estado;

                                fechaHora = fechaHora.AddMinutes(15);
                                f++;
                                workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                workSheet.Cells[f, 3].Value = fechaHora.Date;
                                workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 1];
                                workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 1] / 4;
                                workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";


                                if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                    (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                {
                                    workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                        Convert.ToInt32(list_cc[x].fa[p]));
                                }

                                workSheet.Cells[f, 10].Value = list_cc[x].estado;

                                fechaHora = fechaHora.AddMinutes(15);
                                f++;
                                workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                                workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                                workSheet.Cells[f, 3].Value = fechaHora.Date;
                                workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                                workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4)];
                                workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                                workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4)] / 4;
                                workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                                fechaHora = fechaHora.AddMinutes(15);

                                if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                                    (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                                {
                                    workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                                        Convert.ToInt32(list_cc[x].fa[p]));
                                }

                                workSheet.Cells[f, 10].Value = list_cc[x].estado;
                            }

                            #endregion
                        }
                    }

                    var allCells = workSheet.Cells[1, 1, f, 11];
                    allCells.AutoFitColumns();
                    excelPackage.Save();

                    Console.WriteLine("Comprimiendo archivo ...");
                    zip.ComprimirArchivo(file.FullName, file.FullName.Replace(".xlsx", ".zip"));
                    file.Delete();

                    // GeneraNuevoCorreo(mail);
                    this.RespondeMail(mail, TipoRespuestaCorreo.ConCurvas);
                    // mail.Move(outbox);


                }
                else if (!hayError)// Caso en el que no se han encontrado datos
                {
                    // this.RespondeCorreo(mail, "No existen curvas facturadas para los cups y periodos solicitados.");
                    this.RespondeMail(mail, TipoRespuestaCorreo.SinCurvas);

                }

            }
            catch (Exception e)
            {
                ficheroLog.AddError("CurvasNetezza.CurvasCuartoHorarias: "
                    + e.Message + " --> " + mail.From.Address
                    + " Fila: " + f + "FechaHora: " + fechaHora);
            }
        }

        private void RespondeMail(EmailMessage mail, TipoRespuestaCorreo respuesta)
        {

            string[] listaArchivos;
            FileInfo file;            
            string enCopia = "";
            string para = "";
            bool firstOnly = true;
            string body = "";
            try
            {
                
                // Comprobamos que tengamos adjuntos para crear el correo
                listaArchivos = Directory.GetFiles(regla.rutaSalvadoAdjuntos,
                    param.GetValue("prefijo_archivo_curvas", DateTime.Now, DateTime.Now));

                switch (respuesta)
                {
                    case TipoRespuestaCorreo.ConCurvas:
                        body = CuerpoMailCustom("html_body");
                        break;
                    case TipoRespuestaCorreo.SinCurvas:
                        body = CuerpoMailCustom("html_body_sin_datos");
                        break;
                    default:
                        body = CuerpoMailCustom("html_body_error_formato");
                        break;
                }

                bool replyToAll = true;
                mail.Reply(body, replyToAll);


                for (int i = 0; i < listaArchivos.Count(); i++)
                {
                    mail.Attachments.AddFileAttachment(listaArchivos[i]);
                }

                mail.Attachments.AddFileAttachment(System.Environment.CurrentDirectory + param.GetValue("mail_image_path", DateTime.Now, DateTime.Now));
                if (param.GetValue("adjuntar_correo_ejemplo", DateTime.Now, DateTime.Now) == "S")
                {
                    mail.Attachments.AddFileAttachment(System.Environment.CurrentDirectory + param.GetValue("correo_ejemplo", DateTime.Now, DateTime.Now));
                }

                ficheroLog.Add("Enviando correo Para: " + para + " CC:" + enCopia);

                if (paramUser.GetValue("enviar_mail", DateTime.Now, DateTime.Now) == "N")
                    mail.Save();
                else
                    mail.Send();

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

        public bool GetCurva(string cups13, DateTime fd, DateTime fh, string estado)
        {
            string strSql = "";
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;

            int numeroPeriodos = 0;
            int year = 0;
            int month = 0;
            int day = 0;
            int j = 0;
            int z = 0;
            DateTime fechaHora = new DateTime();
            utilidades.Fechas fechas = new utilidades.Fechas();

            try
            {
                strSql = "SELECT CD_PUNTO_MED as CUPS15, CD_CUPS_EXT as CUPS22, FH_LECT_REGISTRO as FECHA,";
                // HORA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_AC_H" + i + " as A" + i + ",";
                }
                // HORA REACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_R1_H" + i + " as R" + i + ",";
                }
                // FUENTE HORA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " CD_FUENTE_HOR_AC_H" + i + " as FH" + i + ",";
                }
                // CUARTOHORARIA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    for (int y = 1; y <= 4; y++)
                    {
                        z++;
                        strSql += " NM_POT_AC_H" + i + "_CUAD" + y + " as V" + z + ",";
                    }
                }
                // FUENTE CUARTOHORARIA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " CD_FUENTE_CUARTH_AC_H" + i + " as FCH" + i + ",";
                }

                if (estado == "F")
                    strSql += " CD_SEC_RESUMEN AS VERSION, CD_ESTADO_CURVA"
                    + " FROM METRA_OWNER.T_ED_H_CURVAS WHERE"
                    + " CD_PUNTO_MED like '" + cups13 + "%' AND"
                    + " (FH_LECT_REGISTRO >= " + fd.ToString("yyyyMMdd")
                    + " AND FH_LECT_REGISTRO <= " + fh.ToString("yyyyMMdd") + ")"
                    + " and CD_ESTADO_CURVA = '" + estado + "'"
                    + " ORDER BY CD_PUNTO_MED, FH_LECT_REGISTRO;";
                else
                    strSql += " CD_SEC_RESUMEN AS VERSION, CD_ESTADO_CURVA"
                    + " FROM METRA_OWNER.T_ED_H_CURVAS WHERE"
                    + " CD_PUNTO_MED = '" + cups13 + "' AND"
                    + " (FH_LECT_REGISTRO >= " + fd.ToString("yyyyMMdd")
                    + " AND FH_LECT_REGISTRO <= " + fh.ToString("yyyyMMdd") + ")"
                    + " and CD_ESTADO_CURVA = '" + estado + "'"
                    + " ORDER BY CD_PUNTO_MED, FH_LECT_REGISTRO;";

                ficheroLog.Add(strSql);

                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    year = Convert.ToInt32(r["FECHA"].ToString().Substring(0, 4));
                    month = Convert.ToInt32(r["FECHA"].ToString().Substring(4, 2));
                    day = Convert.ToInt32(r["FECHA"].ToString().Substring(6, 2));

                    fechaHora = new DateTime(year, month, day, 0, 0, 0);
                    numeroPeriodos = fechas.NumPeriodosHorarios(fechaHora);

                    EndesaEntity.medida.CurvaCuartoHorariaInformes c = new EndesaEntity.medida.CurvaCuartoHorariaInformes();

                    if (r["CUPS15"] != System.DBNull.Value)
                        c.cups15 = r["CUPS15"].ToString();

                    if (r["CUPS22"] != System.DBNull.Value)
                        c.cups22 = r["CUPS22"].ToString();

                    if (r["CD_ESTADO_CURVA"] != System.DBNull.Value)
                        c.estado = r["CD_ESTADO_CURVA"].ToString() == "F" ? "FACTURADA" : "REGISTRADA";

                    c.fecha = fechaHora;
                    c.numPeriodos = numeroPeriodos;

                    j = 0;
                    for (int h = 1; h <= 25; h++)
                    {
                        if (r["A" + h] != System.DBNull.Value && r["A" + h].ToString() != "")
                            c.a[h] = Convert.ToDouble(r["A" + h]);

                        if (r["FH" + h] != System.DBNull.Value && r["FH" + h].ToString().Trim() != "")
                            c.fa[h] = r["FH" + h].ToString();

                        if (r["R" + h] != System.DBNull.Value && r["R" + h].ToString() != "")
                            c.r[h] = Convert.ToDouble(r["R" + h]);

                        if (r["FCH" + h] != System.DBNull.Value && r["FCH" + h].ToString().Trim() != "")
                            c.fc[h] = r["FCH" + h].ToString();
                    }
                    for (int cc = 1; cc <= 100; cc++)
                        if (r["V" + cc] != System.DBNull.Value && r["V" + cc].ToString() != "")
                            c.value[cc] = Convert.ToDouble(r["V" + cc]);

                    // ficheroLog.Add("Añadiendo a lista --> list_cc" + c.cups15 + " - " + fechaHora);
                    if (!list_cc.Exists(zz => zz.cups15 == c.cups15 && zz.fecha == c.fecha))
                        list_cc.Add(c);


                }

                db.CloseConnection();
                return false;

            }
            catch (Exception e)
            {
                ficheroLog.AddError("CurvasRedShift.GetCurva: " + e.Message);
                return true;
            }

        }

        private FolderId GetFolderID(string folderName)
        {
            FolderId fid = new FolderId(WellKnownFolderName.MsgFolderRoot, mb);
            Folder rootfolder = Folder.Bind(service, fid);
            rootfolder.Load();

            foreach (Folder folder in rootfolder.FindFolders(new FolderView(100)))
            {
                // Finds the emails in a certain folder, in this case the Junk Email

                // Console.WriteLine(folder.DisplayName);
                // This IF limits what folder the program will seek
                if (folder.DisplayName == folderName)
                {
                    // Trust me, the ID is a pain if you want to manually copy and paste it. This stores it in a variable
                    return folder.Id;

                }
            }

            return null;
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

    }
}
