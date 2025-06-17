using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Xml;
using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System.Threading;

namespace EndesaBusiness.medida
{
    public class BuzonPNT
    {
        ExchangeService service;
        Mailbox mb;
        FolderId fid;
        FolderId outboxNoNuestros;
        FolderId outboxNuestros;
        Folder inbox;
        Folder outbox;
        EndesaEntity.ReglaCorreo regla;

        utilidades.Param param;
        utilidades.ParamUser paramUser;

        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_pnts");

        Dictionary<string, EndesaEntity.MeterPoint> mpListCUPS13;
        Dictionary<string, EndesaEntity.MeterPoint> mpListCUPS20;
        public BuzonPNT()
        {
            //utilidades.Credenciales credenciales = new utilidades.Credenciales();
            utilidades.Credenciales credenciales = new utilidades.Credenciales("RSIOPEGMA001");

            mpListCUPS13 = new Dictionary<string, EndesaEntity.MeterPoint>();
            mpListCUPS20 = new Dictionary<string, EndesaEntity.MeterPoint>();

            param = new utilidades.Param("pnt_param", servidores.MySQLDB.Esquemas.MED);            
            paramUser = new utilidades.ParamUser("pnt_param_user", "RSIOPEGMA001", servidores.MySQLDB.Esquemas.MED);
            regla = new EndesaEntity.ReglaCorreo();

            service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials(credenciales.server_user, credenciales.server_password);
            //service.AutodiscoverUrl(credenciales.server_user, RedirectionUrlValidationCallback);
            service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            //service.Url = new Uri("smtp.outlook.com");
            //mb = new Mailbox(paramUser.GetValue("buzon", DateTime.Now, DateTime.Now));
            mb = new Mailbox("eepnt@enel.com");
            fid = new FolderId(WellKnownFolderName.Inbox, mb);
            inbox = Folder.Bind(service, fid);


            outboxNoNuestros = GetFolderID(param.GetValue("carpeta_buzon_no_nuestros", DateTime.Now, DateTime.Now));
            outboxNuestros = GetFolderID(param.GetValue("carpeta_buzon_destino", DateTime.Now, DateTime.Now));

            regla.asunto = param.GetValue("asunto", DateTime.Now, DateTime.Now);

            regla.moverCorreo = true;
            regla.carpetaDestinoCorreo = param.GetValue("carpeta_buzon_destino", DateTime.Now, DateTime.Now);
            regla.ignorar_asunto = param.GetValue("ignorar_asunto", DateTime.Now, DateTime.Now) == "S";
            regla.de = param.GetValue("de", DateTime.Now, DateTime.Now);

            regla.buzon_adicional = paramUser.GetValue("buzon_adicional", DateTime.Now, DateTime.Now) == "S";
            regla.buzon = paramUser.GetValue("buzon", DateTime.Now, DateTime.Now);
            regla.carpeta = paramUser.GetValue("buzon_carpeta", DateTime.Now, DateTime.Now);
            regla.conAdjuntos = true;
            regla.guardarAdjuntos = true;
            regla.carpetaDestinoAlternativo = param.GetValue("carpeta_buzon_no_nuestros", DateTime.Now, DateTime.Now);
            DirectoryInfo dir = new DirectoryInfo(param.GetValue("inbox", DateTime.Now, DateTime.Now));
            if (!dir.Exists)
                dir.Create();

            regla.rutaSalvadoAdjuntos = param.GetValue("inbox", DateTime.Now, DateTime.Now);

            GetInventoryMeterPoint();

        }


        public void BorrarPapelera()
        {
            if (inbox != null)
            {
                FindItemsResults<Item> items = inbox.FindItems(new ItemView(300));

                foreach (EmailMessage eMail in items)
                {
                    eMail.Delete(DeleteMode.HardDelete);
                }
            }
        }


        public void RecorreInbox()
        {
            EmailMessage newEmail = null;
            bool esNuestro = false;
            List<EndesaEntity.Correo> lista_expedientes_gp = new List<EndesaEntity.Correo>();
            List<EndesaEntity.Correo> lista_expedientes_op = new List<EndesaEntity.Correo>();
            string[] listaArchivos;
            FileInfo file;

            if (inbox != null)
            {

                // SearchFilter SF = new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false);

                //List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
                //searchFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.IsSubmitted, "Test"));

                ItemView view = new ItemView(1000, 0, OffsetBasePoint.Beginning);
                view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                view.PropertySet.Add(ItemSchema.DateTimeReceived);
                view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);


                //FindItemsResults<Item> items = inbox.FindItems(new ItemView(1000));
                FindItemsResults<Item> items = inbox.FindItems(view);

                foreach (EmailMessage eMail in items)
                {


                    newEmail = EmailMessage.Bind(service, eMail.Id, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments, ItemSchema.HasAttachments));
                    newEmail.Load();

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
                                if (attachment.Name.Contains(".xml"))
                                {
                                    string filePath = Path.Combine(regla.rutaSalvadoAdjuntos, attachment.Name);
                                    FileAttachment fileAttachment = attachment as FileAttachment;
                                    fileAttachment.Load(filePath);
                                }


                            }

                            esNuestro = MiraDentroArchivoXML();
                            if (esNuestro)
                            {
                                //expendientesNuestros++;
                                lista_expedientes_op.Add(c);


                                listaArchivos = Directory.GetFiles(regla.rutaSalvadoAdjuntos);
                                for (int i = 0; i < listaArchivos.Count(); i++)
                                {
                                    file = new FileInfo(listaArchivos[i]);
                                    file.Delete();
                                }

                                GeneraNuevoCorreo(newEmail);
                                Thread.Sleep(500);
                                newEmail.Move(outboxNuestros);
                            }
                            else
                            {

                                lista_expedientes_gp.Add(c);
                                //Thread.Sleep(500);
                                listaArchivos = Directory.GetFiles(regla.rutaSalvadoAdjuntos);
                                for (int i = 0; i < listaArchivos.Count(); i++)
                                {
                                    file = new FileInfo(listaArchivos[i]);
                                    file.Delete();
                                }
                                newEmail.Move(outboxNoNuestros);
                                Thread.Sleep(500);

                            }
                        }
                    }
                }

                //Console.WriteLine("");
                //Console.WriteLine("");
                //Console.WriteLine("********************************");
                //Console.WriteLine("********************************");
                //Console.WriteLine("****** Moviendo mensajes *******");
                //Console.WriteLine("********************************");
                //Console.WriteLine("********************************");
                //Console.WriteLine("");
                //Console.WriteLine("");


                //ItemView view2 = new ItemView(1000, 0, OffsetBasePoint.Beginning);
                //view2.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                //view2.PropertySet.Add(ItemSchema.DateTimeReceived);
                //view2.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);


                ////FindItemsResults<Item> items = inbox.FindItems(new ItemView(1000));
                //FindItemsResults<Item> items2 = inbox.FindItems(view2);

                //foreach (EmailMessage eMail in items2)
                //{


                //    newEmail = EmailMessage.Bind(service, eMail.Id, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments, ItemSchema.HasAttachments));
                //    newEmail.Load();

                //    Console.CursorLeft = 0;
                //    Console.Write(newEmail.DateTimeReceived.ToString("dd/MM/yyyy HH:mm:ss") + " " + newEmail.Sender.Address + " - " + newEmail.Subject);


                //    if (lista_expedientes_gp.Count == 0 && lista_expedientes_op.Count == 0)
                //        break;

                //    if (newEmail != null)
                //    {
                //        if (newEmail.Subject != null)
                //        {

                //            Console.CursorLeft = 0;
                //            Console.Write(newEmail.DateTimeReceived.ToString("dd/MM/yyyy HH:mm:ss"));

                //            for (int i = 0; i < lista_expedientes_gp.Count; i++)
                //            {



                //                if (newEmail.Sender.Address == lista_expedientes_gp[i].sender &&
                //                    newEmail.Subject == lista_expedientes_gp[i].subject &&
                //                    newEmail.DateTimeReceived == lista_expedientes_gp[i].receivedTime)
                //                {

                //                    newEmail.Move(outboxNoNuestros);
                //                    lista_expedientes_gp.RemoveAt(i);
                //                    break;
                //                }

                //            }

                //            for (int i = 0; i < lista_expedientes_op.Count; i++)
                //            {
                //                if (newEmail.Sender.Address == lista_expedientes_gp[i].sender &&
                //                    newEmail.Subject == lista_expedientes_gp[i].subject &&
                //                    newEmail.DateTimeReceived == lista_expedientes_gp[i].receivedTime)
                //                {
                //                    newEmail.Move(outboxNuestros);
                //                    lista_expedientes_op.RemoveAt(i);
                //                    break;
                //                }

                //            }

                //        }


                //    }

                //}

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

        private void GetAttachmentsFromEmail(ExchangeService service, ItemId itemId)
        {
            // Bind to an existing message item and retrieve the attachments collection.
            // This method results in an GetItem call to EWS.
            EmailMessage message = EmailMessage.Bind(service, itemId, new PropertySet(ItemSchema.Attachments));

            // Iterate through the attachments collection and load each attachment.
            foreach (Attachment attachment in message.Attachments)
            {
                if (attachment is FileAttachment)
                {
                    FileAttachment fileAttachment = attachment as FileAttachment;

                    // Load the attachment into a file.
                    // This call results in a GetAttachment call to EWS.
                    fileAttachment.Load("C:\\temp\\pnts\\" + fileAttachment.Name);

                    Console.WriteLine("File attachment name: " + fileAttachment.Name);
                }
                else // Attachment is an item attachment.
                {
                    ItemAttachment itemAttachment = attachment as ItemAttachment;

                    // Load attachment into memory and write out the subject.
                    // This does not save the file like it does with a file attachment.
                    // This call results in a GetAttachment call to EWS.
                    itemAttachment.Load();

                    Console.WriteLine("Item attachment name: " + itemAttachment.Name);
                }
            }
        }

        private bool MiraDentroArchivoXML()
        {
            string cod_ini = "";
            string cod_fin = "";
            string valor = "";
            string inbox = "";
            string[] listaArchivos;
            bool cupsEsNuestro = false;
            string archivo = "";

            List<EndesaEntity.medida.PNT_XML> list = new List<EndesaEntity.medida.PNT_XML>();

            try
            {
                inbox = param.GetValue("inbox", DateTime.Now, DateTime.Now);

                listaArchivos = Directory.GetFiles(regla.rutaSalvadoAdjuntos, "*.xml");
                for (int i = 0; i < listaArchivos.Count(); i++)
                {

                    EndesaEntity.medida.PNT_XML xml = new EndesaEntity.medida.PNT_XML();
                    EndesaEntity.medida.PNT_XML_Detail xml_detail = new EndesaEntity.medida.PNT_XML_Detail();
                    FileInfo file = new FileInfo(listaArchivos[i]);

                    xml.archivo = file.Name;
                    archivo = xml.archivo;

                    XmlTextReader r = new XmlTextReader(file.FullName);
                    while (r.Read())
                    {

                        switch (r.NodeType)
                        {
                            case XmlNodeType.Element: // The node is an element.
                                cod_ini = r.Name;
                                break;
                            case XmlNodeType.Text: //Display the text in each element.
                                valor = utilidades.FuncionesTexto.ArreglaAcentos(r.Value);
                                break;
                            case XmlNodeType.EndElement: //Display the end of the element.
                                cod_fin = r.Name;
                                break;
                        }

                        #region XML

                        if (cod_ini == cod_fin)
                            switch (cod_ini)
                            {
                                case "CodigoDelProceso":
                                    xml.codigoDelProceso = valor;
                                    break;
                                case "CodigoDePaso":
                                    xml.codigoDePaso = valor;
                                    break;
                                case "CodigoDeSolicitud":
                                    xml.codigoDeSolicitud = Convert.ToInt32(valor);
                                    break;
                                case "NumExpediente":
                                    xml.numExpediente = valor;
                                    break;
                                case "Codigo":
                                    xml.cups = valor;
                                    break;
                                case "FechaSolicitud":
                                    xml.fechaSolicitud = Convert.ToDateTime(valor.Substring(0, 10) + " " + valor.Substring(11, 8));
                                    break;
                                case "FechaInspeccion":
                                    xml.fechaInspeccion = Convert.ToDateTime(valor.Substring(0, 10));
                                    break;
                                case "FechaAltaExpediente":
                                    xml.fechaAltaExpediente = Convert.ToDateTime(valor.Substring(0, 10));
                                    break;
                                case "Comentarios":
                                    xml.comentarios = valor;
                                    break;
                                case "Descripcion":
                                    xml.descripcion = valor;
                                    break;
                                case "FactorCorreccion":
                                    xml.factorCorreccion = valor;
                                    break;
                                case "FechaNormalizacion":
                                    xml.fechaNormalizacion = Convert.ToDateTime(valor.Substring(0, 10));
                                    break;

                            }



                        if (cod_ini == cod_fin)
                        {
                            if (cod_fin.Contains("PotenciaAFacturarP"))
                                xml_detail.potenciaAFacturar[Convert.ToInt32(cod_ini.Replace("PotenciaAFacturarP", "")) - 1] = Convert.ToInt32(valor);
                            if (cod_fin.Contains("EnergiaAFacturarP"))
                                xml_detail.energiaAFacturar[Convert.ToInt32(cod_ini.Replace("EnergiaAFacturarP", "")) - 1] = Convert.ToInt32(valor);
                            if (cod_fin.Contains("ReactivaAFacturarP"))
                                xml_detail.reactivaAFacturar[Convert.ToInt32(cod_ini.Replace("ReactivaAFacturarP", "")) - 1] = Convert.ToInt32(valor);
                            if (cod_fin.Contains("ExcesoPotAFacturarP"))
                                xml_detail.excesoPotAFacturar[Convert.ToInt32(cod_ini.Replace("ExcesoPotAFacturarP", "")) - 1] = Convert.ToInt32(valor);
                        }

                        if (cod_fin == "FechaInicioValoracion" && cod_ini == "FechaInicioValoracion")
                            xml_detail.fechaInicioValoracion = Convert.ToDateTime(valor.Substring(0, 10));
                        if (cod_fin == "FechaFinValoracion" && cod_ini == "FechaFinValoracion")
                            xml_detail.fechaFinValoracion = Convert.ToDateTime(valor.Substring(0, 10));


                        if (cod_fin == "RegistroValAnomalias")
                        {
                            xml.lista_potencias.Add(xml_detail);
                            xml_detail = new EndesaEntity.medida.PNT_XML_Detail();
                        }



                        #endregion

                    }

                    r.Close();

                    if (ExistMeterPoint(xml.cups.Substring(0, 20)))
                    {
                        //ficheroLog.Add("CUPS: " + xml.cups + " OPERACIONES");
                        cupsEsNuestro = true;
                        // file.CopyTo(@"\\e20aemsioa00.enelint.global\M\Temp\XML_PNTs\" + file.Name, true);
                        list.Add(xml);

                    }
                    else
                    {
                        //ficheroLog.Add("CUPS: " + xml.cups + " GP");
                        file.Delete();
                    }

                }

                //GeneraExcel(list);
                //GuardadoBBDD(list);

                return cupsEsNuestro;
            }
            catch (System.Exception e)
            {

                ficheroLog.AddError("MiraDentroArchivoXML: " + e.Message);
                return cupsEsNuestro;
            }

        }

        private void GuardadoBBDD(List<EndesaEntity.medida.PNT_XML> lista)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            bool firstOnly = true;
            int num_reg = 0;

            try
            {

                foreach (EndesaEntity.medida.PNT_XML xml in lista)
                {
                    if (firstOnly)
                    {
                        strSql = "replace into pnt_xml (codigoDelProceso, codigoDePaso, codigoDeSolicitud," +
                            " numExpediente, cups, fechaSolicitud, fechaInspeccion, fechaAltaExpediente, factorCorreccion,  comentarios, descripcion," +
                            " archivo, f_ult_mod) values ";
                        firstOnly = false;
                    }

                    num_reg++;

                    #region Campos

                    if (xml.codigoDelProceso != null)
                        strSql += "('" + xml.codigoDelProceso + "'";
                    else
                        strSql += "(null";

                    if (xml.codigoDePaso != null)
                        strSql += ", '" + xml.codigoDePaso + "'";
                    else
                        strSql += ", null";

                    if (xml.codigoDeSolicitud > 0)
                        strSql += ", " + xml.codigoDeSolicitud;
                    else
                        strSql += ", null";

                    if (xml.numExpediente != null)
                        strSql += ", '" + xml.numExpediente + "'";
                    else
                        strSql += ", null";

                    if (xml.cups != null)
                        strSql += ", '" + xml.cups + "'";
                    else
                        strSql += ", null";

                    if (xml.fechaSolicitud > DateTime.MinValue)
                        strSql += ", '" + xml.fechaSolicitud.ToString("yyyy-MM-dd") + "'";
                    else
                        strSql += ", null";

                    if (xml.fechaInspeccion > DateTime.MinValue)
                        strSql += ", '" + xml.fechaInspeccion.ToString("yyyy-MM-dd") + "'";
                    else
                        strSql += ", null";

                    if (xml.fechaAltaExpediente > DateTime.MinValue)
                        strSql += ", '" + xml.fechaAltaExpediente.ToString("yyyy-MM-dd") + "'";
                    else
                        strSql += ", null";

                    if (xml.factorCorreccion != "" && xml.factorCorreccion != null)
                        strSql += ", '" + utilidades.FuncionesTexto.ArreglaAcentos(xml.factorCorreccion) + "'";
                    else
                        strSql += ", null";

                    if (xml.comentarios != null)
                        strSql += ", '" + utilidades.FuncionesTexto.ArreglaAcentos(xml.comentarios) + "'";
                    else
                        strSql += ", null";

                    if (xml.descripcion != null)
                        strSql += ", '" + utilidades.FuncionesTexto.ArreglaAcentos(xml.descripcion) + "'";
                    else
                        strSql += ", null";

                    strSql += ", '" + xml.archivo + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'),";
                    #endregion

                    if (num_reg > 250)
                    {
                        db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                        command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        num_reg = 0;
                        strSql = "";
                        firstOnly = true;
                    }

                }

                if (num_reg > 0)
                {
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    num_reg = 0;
                    strSql = "";
                }

            }
            catch (System.Exception e)
            {
                ficheroLog.AddError("FuncionesPNTs.GuardadoBBDD --> " + e.Message);

            }
        }

        public bool ExistMeterPoint(string meterpoint)
        {
            EndesaEntity.MeterPoint a;
            if (meterpoint.Length == 13)
                return mpListCUPS13.TryGetValue(meterpoint, out a);
            else
                return mpListCUPS20.TryGetValue(meterpoint, out a);
        }

        private void GetInventoryMeterPoint()
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql = "";



            strSql = "select p.CUPS13, p.CUPS20 from med_listado_scea_cups_historico p group by p.CUPS13;";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                EndesaEntity.MeterPoint c = new EndesaEntity.MeterPoint();

                c.meterPointSort = reader["CUPS13"].ToString();
                c.meterPointLong = reader["CUPS20"].ToString();

                EndesaEntity.MeterPoint a;
                if (!mpListCUPS13.TryGetValue(c.meterPointSort, out a))
                    mpListCUPS13.Add(c.meterPointSort, c);
                EndesaEntity.MeterPoint b;
                if (!mpListCUPS20.TryGetValue(c.meterPointLong, out b))
                    mpListCUPS20.Add(c.meterPointLong, c);
            }
        }

        private void GeneraNuevoCorreo(EmailMessage mail)
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

            body = "Fecha de creacción: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "<p></p>"
                        + "<p></p>"
                        + mail.Body;

            from = param.GetValue("sender", DateTime.Now, DateTime.Now);
            to = param.GetValue("sender", DateTime.Now, DateTime.Now);
            subject = param.GetValue("prefijo_asunto", DateTime.Now, DateTime.Now) + " " + mail.Subject;


            // EndesaBusiness.mail.MailExchangeServer mes = new EndesaBusiness.mail.MailExchangeServer("RSIOPEGMA001");
            EndesaBusiness.office365.OAuth_Mail mes = new EndesaBusiness.office365.OAuth_Mail();
            foreach (Attachment attachment in mail.Attachments)
            {


                string filePath = Path.Combine(regla.rutaSalvadoAdjuntos, attachment.Name);
                FileAttachment fileAttachment = attachment as FileAttachment;
                fileAttachment.Load(filePath);
                if (firstOnly)
                {
                    adjuntos = filePath;
                    firstOnly = false;
                }

                else
                    adjuntos += ";" + filePath;


            }

            mes.SendMail(from, to, cc, subject, body, adjuntos);
            Thread.Sleep(1000);
            mes = null;
            listaArchivos = Directory.GetFiles(regla.rutaSalvadoAdjuntos);

            for (int i = 0; i < listaArchivos.Count(); i++)
            {
                file = new FileInfo(listaArchivos[i]);
                file.Delete();
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
    }
}
