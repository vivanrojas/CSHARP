using EndesaBusiness.logs;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace GestionOperaciones.forms.contratacion
{
    public partial class FrmEndesaEnergia21 : Form
    {
        EndesaBusiness.contratacion.eexxi.EEXXI xxi;
        EndesaBusiness.contratacion.eexxi.Inventario inventario;
        EndesaBusiness.contratacion.eexxi.Casos casos;
        EndesaBusiness.utilidades.Param p;
        EndesaBusiness.logs.Log ficheroLog;

        DataGridViewCell cups20;
        DataGridViewCell codigo_solicitud;
        EndesaBusiness.contratacion.eexxi.EstadosCasos estados_casos;
        EndesaBusiness.contratacion.eexxi.Facturacion facturas;

        int newSortColumn;
        ListSortDirection newColumnDirection = ListSortDirection.Ascending;
        EndesaBusiness.global.Provincias provincias;
        EndesaBusiness.global.Municipios municipio;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        bool habilitada_traza = false;

        public FrmEndesaEnergia21()
        {

            usage.Start("Contratación", "FrmEndesaEnergia21" ,"N/A");
            InitializeComponent();

            ficheroLog = new EndesaBusiness.logs.Log(Environment.CurrentDirectory, "logs", "EndesaEnergia21");
            p = new EndesaBusiness.utilidades.Param("eexxi_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
            habilitada_traza = p.GetValue("habilitar_traza") == "S";


            //if(System.Environment.UserName.ToUpper() == "TROM001" || 
            //    System.Environment.UserName.ToUpper() == "ES52369779K" ||
            //    System.Environment.UserName.ToUpper() == "ES02255021D")
            //{
            //    this.pruebaToolStripMenuItem.Visible = true;
            //}else
            //{
            //    this.pruebaToolStripMenuItem.Visible = false;
            //}
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CargaXML()
        {

            double percent = 0;
            int f_avisos = 0;
            int f_altas = 0;
            int f_bajas = 0;
            int f_duplicados = 0;
            int f_errores = 0;
            int f_otros = 0;
            int f_no_endesa = 0;
            int total_archivos = 0;
            DialogResult result = DialogResult.Yes;
            int ii = 0;

            EndesaBusiness.contratacion.eexxi.Aviso_a_COR aviso_cor =
                new EndesaBusiness.contratacion.eexxi.Aviso_a_COR();
            bool error_xml = false;

            EndesaEntity.contratacion.xxi.XML_Datos xml;

            List<EndesaEntity.contratacion.xxi.XML_Datos> lista_altas = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
            List<EndesaEntity.contratacion.xxi.XML_Datos> lista_bajas = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
            List<EndesaEntity.contratacion.xxi.XML_Datos> lista_t101 = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
            List<EndesaEntity.contratacion.xxi.XML_Datos> lista_t101_paso_a_cor = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
            List<EndesaEntity.contratacion.xxi.XML_Datos> lista_duplicados = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
            List<EndesaEntity.contratacion.xxi.XML_Datos> lista_Bajas_definitiva_mySQL =
                            new List<EndesaEntity.contratacion.xxi.XML_Datos>();

            EndesaBusiness.cups.TarifaATR tarifaATR = new EndesaBusiness.cups.TarifaATR();
            EndesaBusiness.contratacion.eexxi.SolicitudesCodigos sol_codigos = new EndesaBusiness.contratacion.eexxi.SolicitudesCodigos();
            EndesaBusiness.contratacion.eexxi.SolicitudesCodigosArchivo sol_codigos_archivo =
                new EndesaBusiness.contratacion.eexxi.SolicitudesCodigosArchivo();

            Dictionary<string, List<EndesaEntity.contratacion.xxi.XML_Min>> dic_inventario_cargas;
            try
            {




                OpenFileDialog d = new OpenFileDialog();
                d.Title = p.GetValue("mensaje_ventana_xml", DateTime.Now, DateTime.Now);
                d.Filter = "zip files|*.zip";
                d.Multiselect = true;

                if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // Primero contamos que están todos los archivos
                    if (d.FileNames.Count() != 5)
                    {

                        result = MessageBox.Show("!!!!Aviso!!!!"
                        + System.Environment.NewLine
                        + "Únicamente ha seleccionado " + d.FileNames.Count() + " de 5 archivos zip."
                        + System.Environment.NewLine
                        + "Para asegurar la calidad del proceso es necesario que seleccione los "
                        + "siguientes tipos de archivo:"
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "Archivo 1: Tipos T1-01"
                        + System.Environment.NewLine
                        + "Archivo 2: Tipos C1-06"
                        + System.Environment.NewLine
                        + "Archivo 3: Tipos C2-06"
                        + System.Environment.NewLine
                        + "Archivo 4: Tipos T1-05"
                        + System.Environment.NewLine
                        + "Archivo 5: Tipos B1-05"
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "A pesar del aviso:"
                        + System.Environment.NewLine
                        + "¿Está complementamente seguro que desea continuar con el proceso?",
                       "Importación ficheros XML",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    }


                    if (result == DialogResult.Yes)
                    {
                        forms.FrmProgressBar pb = new FrmProgressBar();

                        if (!Directory.Exists(p.GetValue("inbox", DateTime.Now, DateTime.Now)))
                            Directory.CreateDirectory(p.GetValue("inbox", DateTime.Now, DateTime.Now));

                        if (!Directory.Exists(p.GetValue("RutaSalidaXML_ERROR")))
                            Directory.CreateDirectory(p.GetValue("RutaSalidaXML_ERROR"));

                        if (!Directory.Exists(p.GetValue("RutaSalidaXML_NO_ENDESA")))
                            Directory.CreateDirectory(p.GetValue("RutaSalidaXML_NO_ENDESA"));

                        BorrarContenidoDirectorio(p.GetValue("inbox", DateTime.Now, DateTime.Now));

                        dic_inventario_cargas = xxi.Carga_Entradas();

                        pb.Text = "Descomprimiendo ...";
                        pb.Show();
                        pb.progressBar.Step = 1;
                        pb.progressBar.Maximum = Convert.ToInt32(d.FileNames.Count());

                        foreach (string fileName in d.FileNames)
                        {
                            ii++;
                            percent = (ii / Convert.ToDouble(d.FileNames.Count())) * 100;
                            pb.progressBar.Increment(1);
                            pb.txtDescripcion.Text = "Extrayendo " + ii.ToString("#,##0") + " / " + d.FileNames.Count().ToString("#,##0") + " archivos --> " + fileName;
                            pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                            pb.Refresh();

                            //EndesaBusiness.utilidades.ZipUnZip zip = new EndesaBusiness.utilidades.ZipUnZip();
                            //zip.Descomprimir(fileName, p.GetValue("inbox", DateTime.Now, DateTime.Now));
                            EndesaBusiness.utilidades.ZIP zip = new EndesaBusiness.utilidades.ZIP();
                            zip.DescomprimirArchivoZip(fileName, p.GetValue("inbox"));

                        }

                        pb.Close();

                        f_altas = 0;
                        string[] files = Directory.GetFiles(p.GetValue("inbox", DateTime.Now, DateTime.Now), "*.xml");
                        total_archivos = files.Count();

                        if (habilitada_traza)
                            ficheroLog.Add("Se han detectado " + total_archivos + " xml");

                        if (Convert.ToDateTime(p.GetValue("fecha_ultima_carga", DateTime.Now, DateTime.Now)).Date < DateTime.Now.Date)
                        {
                            xxi.SolicitudesTMP_a_Solicitudes();
                            xxi.Borra_tabla("eexxi_solicitudes_tmp");
                            p.UpdateParameter("fecha_ultima_carga", DateTime.Now.ToString("yyyy-MM-dd"));
                            p = new EndesaBusiness.utilidades.Param("eexxi_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
                        }

                        EndesaBusiness.contratacion.eexxi.Inventario inv = new EndesaBusiness.contratacion.eexxi.Inventario();

                        Cursor.Current = Cursors.WaitCursor;

                        #region Lee XML
                        pb = new FrmProgressBar();
                        pb.Text = "Carga archivos XML";
                        pb.Show();
                        pb.progressBar.Step = 1;
                        pb.progressBar.Maximum = files.Count();

                        for (int i = 0; i < files.Count(); i++)
                        {
                            FileInfo file = new FileInfo(files[i]);

                            #region ProgressBar
                            percent = (i / Convert.ToDouble(files.Count())) * 100;
                            pb.progressBar.Increment(1);
                            pb.txtDescripcion.Text = "Importando " + i.ToString("#,##0") + " / " + files.Count().ToString("#,##0")
                                + " archivos --> " + file.Name;
                            pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                            pb.Refresh();
                            #endregion

                            xml = TrataXML(file.FullName);

                            if (xml != null)
                            {
                                if (habilitada_traza)
                                {
                                    ficheroLog.Add("leido xml: " + file.Name + " Proceso: " + xml.codigoDelProceso + xml.codigoDePaso);
                                }

                                if (xml.codigoDelProceso == "T1" && xml.codigoDePaso == "05" &&
                                    (ExisteT102(dic_inventario_cargas, xml)))
                                {
                                    MessageBox.Show("Detectado rechazo."
                                       + System.Environment.NewLine
                                       + System.Environment.NewLine
                                       + "Se ha detectado rechazo para el CUPS: "
                                       + xml.cups + " !!!!!!!!!!"
                                       + System.Environment.NewLine
                                       + System.Environment.NewLine
                                       + System.Environment.NewLine
                                       + System.Environment.NewLine,
                                       "Detectado rechazo",
                                       MessageBoxButtons.OK,
                                       MessageBoxIcon.Warning);
                                }



                                if (!ExisteEntrada(dic_inventario_cargas, xml))
                                {

                                    if (Empresa_XXI(xml))
                                    {
                                        if (Es_Alta(xml))
                                        {
                                            f_altas++;

                                            if (tarifaATR.EsTarifaEEXXI(xml.tarifaATR) && (xml.potenciaPeriodo[6] > 50000))
                                            {
                                                dic_inventario_cargas = ADD_dic_inventario_cargas(dic_inventario_cargas, xml);

                                                lista_altas.Add(xml);
                                            }

                                        }
                                        else if (Es_Aviso_Alta(xml))
                                        {
                                            f_avisos++;

                                            if (tarifaATR.EsTarifaEEXXI(xml.tarifaATR) && (xml.potenciaPeriodo[6] > 50000))
                                            {
                                                dic_inventario_cargas = ADD_dic_inventario_cargas(dic_inventario_cargas, xml);

                                                lista_t101_paso_a_cor.Add(xml);
                                                lista_t101.Add(xml);
                                            }

                                        }
                                        else if (Es_Baja(xml))
                                        {

                                            f_bajas++;
                                            lista_bajas.Add(xml);
                                        }
                                        else
                                            f_otros++;
                                    }
                                    else
                                    {
                                        f_no_endesa++;
                                        FileInfo ficheroDestino = new FileInfo(p.GetValue("RutaSalidaXML_NO_ENDESA") + file.Name);
                                        if (ficheroDestino.Exists)
                                            ficheroDestino.Delete();
                                        file.CopyTo(p.GetValue("RutaSalidaXML_NO_ENDESA") + file.Name);
                                    }


                                }
                                else
                                {
                                    lista_duplicados.Add(xml);
                                    f_duplicados++;
                                }

                            }
                            else
                            {
                                f_errores++;
                                error_xml = true;
                                FileInfo ficheroDestino = new FileInfo(p.GetValue("RutaSalidaXML_ERROR") + file.Name);
                                if (ficheroDestino.Exists)
                                    ficheroDestino.Delete();
                                file.CopyTo(p.GetValue("RutaSalidaXML_ERROR") + file.Name);
                            }

                        }
                        pb.Close();
                        #endregion


                        #region AVISO ALTAS



                        //List<EndesaEntity.contratacion.Solicitudes_Codigos_Tabla> lista_avisoAltas =
                        //        sol_codigos_archivo.dic.Where(z => z.Value.descripcion == "AVISO ALTA").Select(z => z.Value).ToList();

                        //for (int z = 0; z < lista_avisoAltas.Count; z++)
                        //{
                        //    pb = new FrmProgressBar();
                        //    files = Directory.GetFiles(p.GetValue("inbox", DateTime.Now, DateTime.Now), "*"
                        //        + lista_avisoAltas[z].codigoproceso + lista_avisoAltas[z].codigopaso
                        //        + "*.xml");

                        //    pb.Text = "Carga archivos XML";
                        //    pb.Show();
                        //    pb.progressBar.Step = 1;
                        //    pb.progressBar.Maximum = files.Count();

                        //    for (int i = 0; i < files.Count(); i++)
                        //    {


                        //        FileInfo file = new FileInfo(files[i]);
                        //        xml = new EndesaEntity.contratacion.xxi.XML_Datos();

                        //        percent = (i / Convert.ToDouble(files.Count())) * 100;
                        //        pb.progressBar.Increment(1);
                        //        pb.txtDescripcion.Text = "Importando AVISO ALTAS T1 01" + lista_avisoAltas[z].codigoproceso + " "
                        //            + i.ToString("#.###") + " / " + files.Count().ToString("#.###") + " archivos --> " + file.Name;
                        //        pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                        //        pb.Refresh();

                        //        if (file.Name.Contains(lista_avisoAltas[z].codigoproceso + lista_avisoAltas[z].codigopaso))
                        //        {
                        //            xml = TrataXML(file.FullName);
                        //            if(xml != null)
                        //            {
                        //                f_avisos++;

                        //                if (tarifaATR.EsTarifaEEXXI(xml.tarifaATR) && (xml.potenciaPeriodo[6] > 50000))
                        //                {
                        //                    //lista_t101_paso_a_cor.Add(TrataXML(file.FullName));
                        //                    lista_t101_paso_a_cor.Add(xml);

                        //                }
                        //                //lista_t101.Add(TrataXML(file.FullName));
                        //                lista_t101.Add(xml);
                        //            }
                        //            else
                        //            {
                        //                error_xml = true;                                    
                        //                file.CopyTo(p.GetValue("RutaSalidaXML_ERROR") + file.Name);
                        //            }

                        //        }
                        //    }

                        //    pb.Close();
                        //}

                        //if(lista_t101_paso_a_cor.Count > 0 && p.GetValue("Aviso_paso_a_COR") == "S")
                        //{
                        //    aviso_cor.GuardadoBBDD_Paso_a_COR(lista_t101_paso_a_cor);
                        //    aviso_cor.GeneraMails_Paso_a_COR(lista_t101_paso_a_cor);

                        //}


                        //if(lista_t101.Count > 0)
                        //{
                        //    xxi.GuardadoBBDD(lista_t101, "eexxi_solicitudes_t101");

                        //}


                        #endregion


                        #region ALTAS

                        //forms.FrmProgressBar pb = new FrmProgressBar();                    

                        //List<EndesaEntity.contratacion.Solicitudes_Codigos_Tabla> lista_altas =
                        //        sol_codigos_archivo.dic.Where(z => z.Value.descripcion == "ALTA").Select(z => z.Value).ToList();

                        //for (int z = 0; z < lista_altas.Count; z++)
                        //{
                        //    pb = new FrmProgressBar();

                        //    files = Directory.GetFiles(p.GetValue("inbox", DateTime.Now, DateTime.Now), "*" 
                        //        + lista_altas[z].codigoproceso + lista_altas[z].codigopaso
                        //        + "*.xml");


                        //    pb.Text = "Carga archivos XML";
                        //    pb.Show();
                        //    pb.progressBar.Step = 1;
                        //    pb.progressBar.Maximum = files.Count();



                        //    for (int i = 0; i < files.Count(); i++)
                        //    {


                        //        FileInfo file = new FileInfo(files[i]);
                        //        xml = new EndesaEntity.contratacion.xxi.XML_Datos();

                        //        percent = (i / Convert.ToDouble(files.Count())) * 100;
                        //        pb.progressBar.Increment(1);
                        //        pb.txtDescripcion.Text = "Importando ALTAS " + lista_altas[z].codigoproceso + " "
                        //            + i.ToString("#.###") + " / " + files.Count().ToString("#.###") + " archivos --> " + file.Name;
                        //        pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                        //        pb.Refresh();


                        //        if (file.Name.Contains(lista_altas[z].codigoproceso + lista_altas[z].codigopaso))
                        //        {
                        //            xml = TrataXML(file.FullName);
                        //            if(xml != null)
                        //            {
                        //                f_altas++;
                        //                //if (xml.cups.Substring(0, 20) == "ES0339000007000005NS")
                        //                //    Console.WriteLine("a");

                        //                if (tarifaATR.EsTarifaEEXXI(xml.tarifaATR) && (xml.potenciaPeriodo[6] > 50000))
                        //                    //lista.Add(TrataXML(file.FullName));
                        //                    lista.Add(xml);
                        //            }
                        //            else
                        //            {
                        //                error_xml = true;
                        //                file.CopyTo(p.GetValue("RutaSalidaXML_ERROR") + file.Name);
                        //            }

                        //        }
                        //    }

                        //    pb.Close();
                        //}

                        //if(p.GetValue("GenerarExcelAltasT105", DateTime.Now,DateTime.Now) == "S")
                        //{
                        //    //xxi.XML_T105_Classic_To_Excel(lista);
                        //} 



                        #endregion

                        //if(lista_altas.Count > 0)
                        //{
                        //    lista = xxi.Completa_T105_con_T101(lista);
                        //    if(lista != null)
                        //        if(lista.Count > 0)
                        //            xxi.GuardadoBBDD(lista, "eexxi_solicitudes_tmp");
                        //    inv.Carga("eexxi_solicitudes_tmp");
                        //    pb.Close();
                        //}


                        #region BAJAS

                        //List<EndesaEntity.contratacion.Solicitudes_Codigos_Tabla> lista_bajas =
                        //        sol_codigos_archivo.dic.Where(z => z.Value.descripcion == "BAJA").Select(z => z.Value).ToList();

                        //for (int z = 0; z < lista_bajas.Count; z++)
                        //{

                        //    pb = new FrmProgressBar();

                        //    files = Directory.GetFiles(p.GetValue("inbox", DateTime.Now, DateTime.Now), "*"
                        //       + lista_bajas[z].codigoproceso + lista_bajas[z].codigopaso
                        //       + "*.xml");



                        //    pb.Text = "Carga archivos XML";
                        //    pb.Show();
                        //    pb.progressBar.Step = 1;
                        //    pb.progressBar.Maximum = files.Count();                        

                        //    for (int i = 0; i < files.Count(); i++)
                        //    {

                        //        FileInfo file = new FileInfo(files[i]);
                        //        xml = new EndesaEntity.contratacion.xxi.XML_Datos();
                        //        EndesaEntity.contratacion.Inventario_Tabla o;

                        //        percent = (i / Convert.ToDouble(files.Count())) * 100;
                        //        pb.progressBar.Increment(1);
                        //        pb.txtDescripcion.Text = "Importando BAJAS " + i.ToString("#.###") + " / " + files.Count().ToString("#.###") + " archivos --> " + file.Name;
                        //        pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                        //        pb.Refresh();


                        //        //if (file.Name.Contains(lista_bajas[z].codigoproceso + lista_bajas[z].codigopaso) &&
                        //        //    inv.dic.TryGetValue(xxi.ExtraeCUPS22_DesdeNombreFichero(file.Name), out o))
                        //        if (file.Name.Contains(lista_bajas[z].codigoproceso + lista_bajas[z].codigopaso))
                        //        {
                        //            xml = TrataXML(file.FullName);
                        //            if(xml != null)
                        //            {
                        //                //if(inv.dic.TryGetValue(xxi.ExtraeCUPS22_DesdeNombreFichero(file.Name), out o)
                        //                f_bajas++;
                        //                //lista.Add(TrataXML(file.FullName));
                        //                lista.Add(xml);
                        //            }
                        //            else
                        //            {
                        //                error_xml = true;
                        //                file.CopyTo(p.GetValue("RutaSalidaXML_ERROR") + file.Name);
                        //            }
                        //        }
                        //    }
                        //    pb.Close();
                        //}




                        #endregion

                        #region Volcan



                        //List<EndesaEntity.contratacion.Solicitudes_Codigos_Tabla> lista_bajas_volcan =
                        //        sol_codigos.dic.Where(z => z.Value.descripcion == "BAJA VOLCAN").Select(z => z.Value).ToList();

                        //Dictionary<string, FileInfo> dic_volcan = new Dictionary<string, FileInfo>();

                        //for (int z = 0; z < lista_bajas_volcan.Count; z++)
                        //{

                        //    files = Directory.GetFiles(p.GetValue("inbox", DateTime.Now, DateTime.Now), "*"
                        //       + lista_bajas_volcan[z].codigoproceso + "_" + lista_bajas_volcan[z].codigopaso
                        //       + "*.xml");



                        //    pb = new FrmProgressBar();
                        //    pb.Text = "Carga archivos XML";
                        //    pb.Show();
                        //    pb.progressBar.Step = 1;
                        //    pb.progressBar.Maximum = files.Count();

                        //    for (int i = 0; i < files.Count(); i++)
                        //    {
                        //        f_bajas++;
                        //        FileInfo file = new FileInfo(files[i]);
                        //        EndesaEntity.contratacion.xxi.XML_Datos xml = new EndesaEntity.contratacion.xxi.XML_Datos();
                        //        EndesaEntity.contratacion.Inventario_Tabla o;

                        //        percent = (i / Convert.ToDouble(files.Count())) * 100;
                        //        pb.progressBar.Increment(1);
                        //        pb.txtDescripcion.Text = "Importando BAJAS VOLCAN " + i.ToString("#.###") + " / " + files.Count().ToString("#.###") + " archivos --> " + file.Name;
                        //        pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                        //        pb.Refresh();

                        //        if (file.Name.Contains(lista_bajas_volcan[z].codigoproceso + "_" + lista_bajas_volcan[z].codigopaso))
                        //        {

                        //            xml = TrataXML(file.FullName);
                        //            if (xml.motivo == "04" && xml.cups.Substring(0, 7) == "ES00316")
                        //            {
                        //                f++;
                        //                c = 1;
                        //                workSheet.Cells[f, c].Value = xml.cups; c++;
                        //                workSheet.Cells[f, c].Value = xml.codigoDeSolicitud; c++;
                        //                workSheet.Cells[f, c].Value = xml.fechaSolicitud.ToString("dd/MM/yyyy"); c++;
                        //                workSheet.Cells[f, c].Value = xml.motivo; c++;
                        //                workSheet.Cells[f, c].Value = xml.fechaActivacion.ToString("dd/MM/yyyy"); c++;



                        //                FileInfo oo;
                        //                if (!dic_volcan.TryGetValue(xml.cups, out oo))

                        //                    dic_volcan.Add(xml.cups.Substring(0, 20), file);

                        //            }

                        //            //lista.Add(TrataXML(file.FullName));


                        //        }



                        //    }
                        //    pb.Close();

                        //    excelPackage.Save();

                        //    //EndesaBusiness.contratacion.PS_AT_HIST ps_at_hist = new EndesaBusiness.contratacion.PS_AT_HIST(dic_volcan.Select(h => h.Key).ToList(),"TODAS");

                        //    //foreach (KeyValuePair<string, FileInfo> p in dic_volcan)
                        //    //{
                        //    //    if(ps_at_hist.ExisteAlta(p.Key))
                        //    //        p.Value.CopyTo(@"c:\Temp\xml_volcan\" + p.Value.Name);
                        //    //}


                        //}

                        #endregion



                        if (lista_altas.Count > 0)
                            xxi.GuardadoBBDD(lista_altas, "eexxi_solicitudes_tmp");

                        if (lista_t101.Count > 0)
                            xxi.GuardadoBBDD(lista_t101, "eexxi_solicitudes_t101");


                        if (lista_duplicados.Count > 0)
                            xxi.GuardadoBBDD(lista_duplicados, "eexxi_solicitudes_duplicadas");

                        if (lista_bajas.Count > 0)
                        {

                            foreach (EndesaEntity.contratacion.xxi.XML_Datos p in lista_bajas)
                            {
                                List<EndesaEntity.contratacion.xxi.XML_Min> o;
                                if (dic_inventario_cargas.TryGetValue(p.cups, out o))
                                {
                                    for (int i = 0; i < o.Count; i++)
                                    {
                                        // Solo consideramos la baja si tenemos
                                        // el paso T105
                                        if (o[i].proceso == "T1" && o[i].paso == "05")
                                            lista_Bajas_definitiva_mySQL.Add(p);
                                    }
                                }

                            }

                            xxi.GuardadoBBDD(lista_Bajas_definitiva_mySQL, "eexxi_solicitudes_tmp");
                        }


                        if (lista_t101_paso_a_cor.Count > 0 && p.GetValue("Aviso_paso_a_COR") == "S")
                        {

                            aviso_cor.GuardadoBBDD_Paso_a_COR(lista_t101_paso_a_cor);
                            aviso_cor.GeneraMails_Paso_a_COR(lista_t101_paso_a_cor);

                            //List<EndesaEntity.contratacion.xxi.XML_Datos> lista_t101_definitiva =
                            //   new List<EndesaEntity.contratacion.xxi.XML_Datos>();
                            //bool encontrado = false;

                            //foreach (EndesaEntity.contratacion.xxi.XML_Datos p in lista_t101_paso_a_cor)
                            //{
                            //    encontrado = false;
                            //    foreach (EndesaEntity.contratacion.xxi.XML_Datos j in lista_Bajas_definitiva_mySQL)
                            //    {

                            //        if (!encontrado && (p.cups == j.cups &&
                            //            j.fechaSolicitud >= p.fechaSolicitud))
                            //            encontrado = true;


                            //    }
                            //    if (!encontrado)
                            //        lista_t101_definitiva.Add(p);
                            //}



                        }


                        if (lista_t101_paso_a_cor.Count > 0 && p.GetValue("Aviso_a_Ventas") == "S")
                        {
                            aviso_cor.GuardadoBBDD_Aviso_Ventas(lista_t101_paso_a_cor);
                            aviso_cor.GeneraMails_Aviso_Ventas();
                        }



                        xxi.ConsultasActualizaDatosSolicitudes();

                        if (!error_xml)
                        {
                            MessageBox.Show("La importación ha concluido correctamente."
                            + System.Environment.NewLine
                            + System.Environment.NewLine
                            + "Se han procesado " + (f_avisos + f_bajas + f_altas + f_duplicados + f_otros).ToString("#,##0") + " archivos de "
                            + total_archivos.ToString("#,##0")
                            + System.Environment.NewLine
                            + System.Environment.NewLine
                            + " Se han encontrado " + f_no_endesa.ToString("#,##0") + " xml de No ENDESA."
                            + System.Environment.NewLine
                            + System.Environment.NewLine,
                            "Importación ficheros XML",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);

                            if (f_altas > 0 || f_bajas > 0)
                                MessageBox.Show("Generar informe de Altas EEXXI."
                                + System.Environment.NewLine
                                + System.Environment.NewLine
                                + "Debe generar el informe de Altas de EEXXI!!!!",
                                "Recordatorio",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                        else
                            MessageBox.Show("La importación ha concluido pero con errores."
                            + System.Environment.NewLine
                            + System.Environment.NewLine
                            + "Se han procesado " + (f_avisos + f_bajas + f_altas + f_duplicados + f_otros).ToString("#,##0")
                            + " archivos de " + total_archivos.ToString("#,##0")
                            + System.Environment.NewLine
                            + System.Environment.NewLine
                            + " Se han encontrado " + f_no_endesa.ToString("#,##0") + " xml de No ENDESA."
                            + System.Environment.NewLine
                            + System.Environment.NewLine
                            + " Se han encontrado " + (f_errores).ToString("#,##0") + " errores."
                            + System.Environment.NewLine
                            + System.Environment.NewLine
                            + "Consulte la ruta " + p.GetValue("RutaSalidaXML_ERROR")
                            + " para ver los XML que no se han podido procesar.",
                            "Importación ficheros XML",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);


                        xxi.ProcesaSolicitudes();
                        LoadData();

                    }
                }
            }
            catch (Exception ex)
            {

                ficheroLog.AddError("FrmEndesaEnergia21.CargaXML: " + ex.Message);

                MessageBox.Show(ex.Message,
                                 "Carga XML",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

            }
        }

        private void CargaXML_SinRestricciones()
        {

            double percent = 0;
            int f_avisos = 0;
            int f_altas = 0;
            int f_bajas = 0;
            int f_duplicados = 0;
            int f_errores = 0;
            int f_otros = 0;
            int total_archivos = 0;
            DialogResult result = DialogResult.Yes;
            int ii = 0;

            EndesaBusiness.contratacion.eexxi.Aviso_a_COR aviso_cor =
                new EndesaBusiness.contratacion.eexxi.Aviso_a_COR();
            bool error_xml = false;

            EndesaEntity.contratacion.xxi.XML_Datos xml;

            List<EndesaEntity.contratacion.xxi.XML_Datos> lista_altas = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
            List<EndesaEntity.contratacion.xxi.XML_Datos> lista_bajas = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
            List<EndesaEntity.contratacion.xxi.XML_Datos> lista_t101 = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
            List<EndesaEntity.contratacion.xxi.XML_Datos> lista_t101_paso_a_cor = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
            List<EndesaEntity.contratacion.xxi.XML_Datos> lista_duplicados = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
            List<EndesaEntity.contratacion.xxi.XML_Datos> lista_Bajas_definitiva_mySQL =
                            new List<EndesaEntity.contratacion.xxi.XML_Datos>();

            EndesaBusiness.cups.TarifaATR tarifaATR = new EndesaBusiness.cups.TarifaATR();
            EndesaBusiness.contratacion.eexxi.SolicitudesCodigos sol_codigos = new EndesaBusiness.contratacion.eexxi.SolicitudesCodigos();
            EndesaBusiness.contratacion.eexxi.SolicitudesCodigosArchivo sol_codigos_archivo =
                new EndesaBusiness.contratacion.eexxi.SolicitudesCodigosArchivo();

            Dictionary<string, List<EndesaEntity.contratacion.xxi.XML_Min>> dic_inventario_cargas;

            OpenFileDialog d = new OpenFileDialog();
            d.Title = p.GetValue("mensaje_ventana_xml", DateTime.Now, DateTime.Now);
            d.Filter = "zip files|*.zip";
            d.Multiselect = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Primero contamos que están todos los archivos
                if (d.FileNames.Count() != 5)
                {

                    result = MessageBox.Show("!!!!Aviso!!!!"
                    + System.Environment.NewLine
                    + "Únicamente ha seleccionado " + d.FileNames.Count() + " de 5 archivos zip."
                    + System.Environment.NewLine
                    + "Para asegurar la calidad del proceso es necesario que seleccione los "
                    + "siguientes tipos de archivo:"
                    + System.Environment.NewLine
                    + System.Environment.NewLine
                    + "Archivo 1: Tipos T1-01"
                    + System.Environment.NewLine
                    + "Archivo 2: Tipos C1-06"
                    + System.Environment.NewLine
                    + "Archivo 3: Tipos C2-06"
                    + System.Environment.NewLine
                    + "Archivo 4: Tipos T1-05"
                    + System.Environment.NewLine
                    + "Archivo 5: Tipos B1-05"
                    + System.Environment.NewLine
                    + System.Environment.NewLine
                    + System.Environment.NewLine
                    + "A pesar del aviso:"
                    + System.Environment.NewLine
                    + "¿Está complementamente seguro que desea continuar con el proceso?",
                   "Importación ficheros XML",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                }


                if (result == DialogResult.Yes)
                {
                    forms.FrmProgressBar pb = new FrmProgressBar();

                    if (!Directory.Exists(p.GetValue("inbox", DateTime.Now, DateTime.Now)))
                        Directory.CreateDirectory(p.GetValue("inbox", DateTime.Now, DateTime.Now));

                    if (!Directory.Exists(p.GetValue("RutaSalidaXML_ERROR")))
                        Directory.CreateDirectory(p.GetValue("RutaSalidaXML_ERROR"));

                    if (!Directory.Exists(p.GetValue("RutaSalidaXML_NO_ENDESA")))
                        Directory.CreateDirectory(p.GetValue("RutaSalidaXML_NO_ENDESA"));


                    BorrarContenidoDirectorio(p.GetValue("inbox", DateTime.Now, DateTime.Now));

                    dic_inventario_cargas = xxi.Carga_Entradas();

                    pb.Text = "Descomprimiendo ...";
                    pb.Show();
                    pb.progressBar.Step = 1;
                    pb.progressBar.Maximum = Convert.ToInt32(d.FileNames.Count());

                    foreach (string fileName in d.FileNames)
                    {
                        ii++;
                        percent = (ii / Convert.ToDouble(d.FileNames.Count())) * 100;
                        pb.progressBar.Increment(1);
                        pb.txtDescripcion.Text = "Extrayendo " + ii.ToString("#.###") + " / " +
                            d.FileNames.Count().ToString("#.###") + " archivos --> " + fileName;
                        pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                        pb.Refresh();

                        EndesaBusiness.utilidades.ZipUnZip zip = new EndesaBusiness.utilidades.ZipUnZip();
                        zip.Descomprimir(fileName, p.GetValue("inbox", DateTime.Now, DateTime.Now));
                        // Descomprimimos en directorio destino inbox
                        //GO.utilidades.ZipUnZip.DescomprimirArchivo(fileName, p.GetValue("inbox", DateTime.Now, DateTime.Now));
                    }

                    pb.Close();

                    f_altas = 0;
                    string[] files = Directory.GetFiles(p.GetValue("inbox", DateTime.Now, DateTime.Now), "*.xml");
                    total_archivos = files.Count();

                    if (Convert.ToDateTime(p.GetValue("fecha_ultima_carga", DateTime.Now, DateTime.Now)).Date < DateTime.Now.Date)
                    {
                        xxi.SolicitudesTMP_a_Solicitudes();
                        xxi.Borra_tabla("eexxi_solicitudes_tmp");
                        p.UpdateParameter("fecha_ultima_carga", DateTime.Now.ToString("yyyy-MM-dd"));
                        p = new EndesaBusiness.utilidades.Param("eexxi_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
                    }

                    EndesaBusiness.contratacion.eexxi.Inventario inv = new EndesaBusiness.contratacion.eexxi.Inventario();

                    Cursor.Current = Cursors.WaitCursor;


                    #region Lee XML
                    pb = new FrmProgressBar();
                    pb.Text = "Carga archivos XML";
                    pb.Show();
                    pb.progressBar.Step = 1;
                    pb.progressBar.Maximum = files.Count();

                    for (int i = 0; i < files.Count(); i++)
                    {
                        FileInfo file = new FileInfo(files[i]);

                        #region ProgressBar
                        percent = (i / Convert.ToDouble(files.Count())) * 100;
                        pb.progressBar.Increment(1);
                        pb.txtDescripcion.Text = "Importando " + i.ToString("#.###") + " / " + files.Count().ToString("#.###")
                            + " archivos --> " + file.Name;
                        pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                        pb.Refresh();
                        #endregion

                        xml = TrataXML(file.FullName);

                        if (xml != null)
                        {
                            if (!ExisteEntrada(dic_inventario_cargas, xml))
                            {

                                if (Es_Alta(xml))
                                {
                                    f_altas++;                                                                                                          
                                    dic_inventario_cargas = ADD_dic_inventario_cargas(dic_inventario_cargas, xml);
                                    lista_altas.Add(xml);

                                }
                                else if (Es_Aviso_Alta(xml))
                                {
                                    f_avisos++;
                                   
                                    dic_inventario_cargas = ADD_dic_inventario_cargas(dic_inventario_cargas, xml);
                                    lista_t101_paso_a_cor.Add(xml);
                                    lista_t101.Add(xml);                                   

                                }
                                else if (Es_Baja(xml))
                                {
                                    f_bajas++;
                                    lista_bajas.Add(xml);
                                }
                                else
                                    f_otros++;
                            }
                            else
                            {
                                lista_duplicados.Add(xml);
                                f_duplicados++;
                            }

                        }
                        else
                        {
                            f_errores++;
                            error_xml = true;
                            FileInfo ficheroDestino = new FileInfo(p.GetValue("RutaSalidaXML_ERROR") + file.Name);
                            if (ficheroDestino.Exists)
                                ficheroDestino.Delete();
                            file.CopyTo(p.GetValue("RutaSalidaXML_ERROR") + file.Name);
                        }

                    }
                    pb.Close();
                    #endregion

                    if (lista_altas.Count > 0)
                        xxi.GuardadoBBDD(lista_altas, "eexxi_solicitudes_tmp");

                    if (lista_t101.Count > 0)
                        xxi.GuardadoBBDD(lista_t101, "eexxi_solicitudes_t101");

                    if (lista_duplicados.Count > 0)
                        xxi.GuardadoBBDD(lista_duplicados, "eexxi_solicitudes_duplicadas");

                    if (lista_bajas.Count > 0)
                    {

                        foreach (EndesaEntity.contratacion.xxi.XML_Datos p in lista_bajas)
                        {
                            List<EndesaEntity.contratacion.xxi.XML_Min> o;
                            if (dic_inventario_cargas.TryGetValue(p.cups, out o))
                            {
                                for (int i = 0; i < o.Count; i++)
                                {
                                    // Solo consideramos la baja si tenemos
                                    // el paso T105
                                    if (o[i].proceso == "T1" && o[i].paso == "05")
                                        lista_Bajas_definitiva_mySQL.Add(p);
                                }
                            }

                        }

                        xxi.GuardadoBBDD(lista_Bajas_definitiva_mySQL, "eexxi_solicitudes_tmp");
                    }


                    if (lista_t101_paso_a_cor.Count > 0 && p.GetValue("Aviso_paso_a_COR") == "S")
                    {

                        aviso_cor.GuardadoBBDD_Paso_a_COR(lista_t101_paso_a_cor);
                        aviso_cor.GeneraMails_Paso_a_COR(lista_t101_paso_a_cor);

                        //List<EndesaEntity.contratacion.xxi.XML_Datos> lista_t101_definitiva =
                        //   new List<EndesaEntity.contratacion.xxi.XML_Datos>();
                        //bool encontrado = false;

                        //foreach (EndesaEntity.contratacion.xxi.XML_Datos p in lista_t101_paso_a_cor)
                        //{
                        //    encontrado = false;
                        //    foreach (EndesaEntity.contratacion.xxi.XML_Datos j in lista_Bajas_definitiva_mySQL)
                        //    {

                        //        if (!encontrado && (p.cups == j.cups &&
                        //            j.fechaSolicitud >= p.fechaSolicitud))
                        //            encontrado = true;


                        //    }
                        //    if (!encontrado)
                        //        lista_t101_definitiva.Add(p);
                        //}



                    }


                    xxi.ConsultasActualizaDatosSolicitudes();

                    if (!error_xml)
                        MessageBox.Show("La importación ha concluido correctamente."
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "Se han procesado " + (f_avisos + f_bajas + f_altas + f_duplicados + f_otros).ToString("#.###") + " archivos de " + total_archivos.ToString("#.###"),
                        "Importación ficheros XML",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    else
                        MessageBox.Show("La importación ha concluido pero con errores."
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "Se han procesado " + (f_avisos + f_bajas + f_altas + f_duplicados + f_otros).ToString("#.###")
                        + " archivos de " + total_archivos.ToString("#.###")
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + " Se han encontrado " + (f_errores).ToString("#.###") + " errores."
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "Consulte la ruta " + p.GetValue("RutaSalidaXML_ERROR")
                        + " para ver los XML que no se han podido procesar.",
                        "Importación ficheros XML",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);


                    xxi.ProcesaSolicitudes();
                    LoadData();

                }
            }
        }

        private Dictionary<string, List<EndesaEntity.contratacion.xxi.XML_Min>> 
            ADD_dic_inventario_cargas(Dictionary<string, List<EndesaEntity.contratacion.xxi.XML_Min>> dic,
            EndesaEntity.contratacion.xxi.XML_Datos xml)
        {

            EndesaEntity.contratacion.xxi.XML_Min c = new EndesaEntity.contratacion.xxi.XML_Min();
            c.proceso = xml.codigoDelProceso;
            c.paso = xml.codigoDePaso;
            c.solicitud = xml.codigoDeSolicitud;
            c.codigo_ree_empresa_emisora = xml.codigoREEEmpresaEmisora;
            c.codigo_ree_empresa_destino = xml.codigoREEEmpresaDestino;

            List<EndesaEntity.contratacion.xxi.XML_Min> o;
            if (!dic.TryGetValue(xml.cups, out o))
            {
                o = new List<EndesaEntity.contratacion.xxi.XML_Min>();
                o.Add(c);
            }
            else
                o.Add(c);
           
            return dic;
        }

        private bool ExisteEntrada(Dictionary<string, List<EndesaEntity.contratacion.xxi.XML_Min>> dic, 
            EndesaEntity.contratacion.xxi.XML_Datos xml)
        {
            bool encontrado = false;
            List<EndesaEntity.contratacion.xxi.XML_Min> o;
            if (dic.TryGetValue(xml.cups, out o))
            {
                foreach (EndesaEntity.contratacion.xxi.XML_Min p in o)
                {
                    if (!encontrado && (p.proceso == xml.codigoDelProceso &&
                         p.paso == xml.codigoDePaso &&
                         p.solicitud == xml.codigoDeSolicitud))
                    {
                        encontrado = true;
                    }
                }
                    
            }
            
            return encontrado;

        }
        
        private bool ExisteT102(Dictionary<string, List<EndesaEntity.contratacion.xxi.XML_Min>> dic,
            EndesaEntity.contratacion.xxi.XML_Datos xml)
        {
            bool encontrado = false;
            List<EndesaEntity.contratacion.xxi.XML_Min> o;
            if (dic.TryGetValue(xml.cups, out o))
            {
                foreach (EndesaEntity.contratacion.xxi.XML_Min p in o)
                {
                    if (!encontrado && (p.proceso == "T1" &&
                         p.paso == "02" && p.solicitud == xml.codigoDeSolicitud
                         && p.codigo_ree_empresa_destino == xml.codigoREEEmpresaEmisora))
                    {
                        encontrado = true;
                    }
                }

            }

            return encontrado;
        }


        private bool Es_Alta(EndesaEntity.contratacion.xxi.XML_Datos xml)
        {
            return (xml.codigoDelProceso == "T1" && xml.codigoDePaso == "05");
        }

        private bool Empresa_XXI(EndesaEntity.contratacion.xxi.XML_Datos xml)
        {
            return (xml.codigoREEEmpresaDestino == p.GetValue("empresa_destino"));
        }

        private bool Es_Aviso_Alta(EndesaEntity.contratacion.xxi.XML_Datos xml)
        {
            return (xml.codigoDelProceso == "T1" && xml.codigoDePaso == "01");
        }

        private bool Es_Baja(EndesaEntity.contratacion.xxi.XML_Datos xml)
        {
            return (xml.codigoDelProceso == "B1" && xml.codigoDePaso == "05")
                || (xml.codigoDelProceso == "B2" && xml.codigoDePaso == "05")
                || (xml.codigoDelProceso == "C1" && xml.codigoDePaso == "06")
                || (xml.codigoDelProceso == "C2" && xml.codigoDePaso == "06");
        }

        private void btn_xml_Click(object sender, EventArgs e)
        {
            CargaXML();           
        }

        private void BorrarContenidoDirectorio(string directorio)
        {
            string[] listaArchivos;
            FileInfo file;

            listaArchivos = Directory.GetFiles(directorio);
            for (int i = 0; i < listaArchivos.Count(); i++)
            {
                file = new FileInfo(listaArchivos[i]);
                file.Delete();
            }
        }

        private EndesaEntity.contratacion.xxi.XML_Datos TrataXML(string fileName)
        {
            string cod_ini = "";
            string cod_fin = "";
            string valor = "";
            bool dentroDireccionCliente = false;
            bool dentroDireccionPS = false;

            int potencia = 0;

            EndesaEntity.contratacion.xxi.XML_Datos xml = new EndesaEntity.contratacion.xxi.XML_Datos();
            FileInfo file = new FileInfo(fileName);

            // XmlDocument

            xml.fichero = file.Name;
            XmlTextReader r;

            try
            {
                r = new XmlTextReader(fileName);
                while (r.Read())
                {
                    
                    switch (r.NodeType)
                    {
                        
                        case XmlNodeType.Element: // The node is an element.
                            cod_ini = r.Name;

                            if (!dentroDireccionCliente)
                                dentroDireccionCliente = (cod_ini == "Cliente");

                            if(!dentroDireccionPS)
                                dentroDireccionPS = (cod_ini == "DireccionPS");

                            if (cod_ini == "Potencia")
                                potencia++;
                            break;
                        case XmlNodeType.Text: //Display the text in each element.
                            valor = EndesaBusiness.utilidades.FuncionesTexto.ArreglaAcentos(r.Value);
                            break;
                        case XmlNodeType.EndElement: //Display the end of the element.
                            cod_fin = r.Name;
                            break;


                    }

                    #region XML


                    if (dentroDireccionCliente)
                    {
                        if (cod_ini == cod_fin)
                        {
                            switch (cod_ini)
                            {
                                case "Pais":                                    
                                    xml.paisCliente = valor;                                    
                                    break;
                                case "Provincia":
                                    xml.provinciaCliente = valor;                                    
                                    break;
                                case "Municipio":                                    
                                    xml.municipioCliente = valor;                                    
                                    break;                                
                                case "TipoVia":                                    
                                    xml.tipoViaCliente = valor;                                    
                                    break;
                                case "CodPostal":                                    
                                    xml.codPostalCliente = valor;                                    
                                    break;
                                case "Calle":                                    
                                    xml.calleCliente = valor;                                    
                                    break;
                                case "NumeroFinca":                                    
                                    xml.numeroFincaCliente = valor;                                    
                                    break;
                                case "AclaradorFinca":
                                    xml.aclaradorFinca = valor;
                                    break;
                                case "Numero":
                                    xml.numero = valor;
                                    break;

                            }

                        }
                    }

                    if (dentroDireccionPS)
                    {
                        dentroDireccionCliente = false;

                        if (cod_ini == cod_fin)
                        {
                            switch (cod_ini)
                            {
                                case "Pais":
                                    xml.pais = valor;
                                    break;
                                case "Provincia":
                                    xml.provincia = valor;
                                    break;
                                case "Municipio":
                                    xml.municipio = valor;
                                    break;
                                case "TipoVia":
                                    xml.tipoVia = valor;
                                    break;
                                case "CodPostal":
                                    xml.codPostal = valor;
                                    break;
                                case "Calle":
                                    xml.calle = valor;
                                    break;
                                case "NumeroFinca":
                                    xml.numeroFinca = valor;
                                    break;
                                case "AclaradorFinca":
                                    xml.aclaradorFinca = valor;
                                    break;                                

                            }

                        }
                    }


                    if (cod_ini == cod_fin)
                        switch (cod_ini)
                        {
                            case "Motivo":
                                xml.motivo = valor;
                                break;
                            case "NombreDePila":
                                xml.razonSocial = valor;
                                break;
                            case "PrimerApellido":
                                xml.razonSocial += " " + valor;
                                break;
                            case "CodigoREEEmpresaEmisora":
                                xml.codigoREEEmpresaEmisora = valor;
                                break;
                            case "CodigoREEEmpresaDestino":
                                xml.codigoREEEmpresaDestino = valor;
                                break;
                            case "CodigoDelProceso":
                                xml.codigoDelProceso = valor;
                                break;
                            case "CodigoDePaso":
                                xml.codigoDePaso = valor;
                                break;
                            case "CodigoDeSolicitud":
                                xml.codigoDeSolicitud = valor;
                                break;  
                            case "SecuencialDeSolicitud":
                                xml.secuencialDeSolicitud = valor;
                                break;
                            case "FechaSolicitud":
                                xml.fechaSolicitud = Convert.ToDateTime(valor.Substring(0, 10) + " " + valor.Substring(11, 8));
                                break;
                            case "CUPS":
                                xml.cups = valor;
                                break;
                            case "CNAE":
                                xml.cnae = valor;
                                break;
                            case "PotenciaExtension":
                                xml.potenciaExtension = Convert.ToDouble(valor);
                                break;
                            case "PotenciaDeAcceso":
                                xml.potenciaDeAcceso = Convert.ToDouble(valor);
                                break;
                            case "PotenciaInstAT":
                                xml.potenciaInstAT = Convert.ToDouble(valor);
                                break;
                            case "IndicativoDeInterrumpibilidad":
                                xml.indicativoDeInterrumpibilidad = valor;
                                break;
                            //case "Pais":
                            //    if (xml.pais != null)
                            //        xml.paisCliente = valor;
                            //    else
                            //        xml.pais = valor;
                            //    break;
                            //case "Provincia":
                            //    if (xml.provincia != null)
                            //        xml.provincia = valor;
                            //    else
                            //        xml.provincia = valor;
                            //    break;
                            //case "Municipio":
                            //    if (xml.municipio != null)
                            //        xml.municipio = valor;
                            //    else
                            //        xml.municipio = valor;
                            //    break;
                            //case "Poblacion":
                            //    if (xml.poblacion != null)
                            //        xml.poblacion = valor;
                            //    else
                            //        xml.poblacion = valor;
                            //    break;
                            //case "DescripcionPoblacion":
                            //    if (xml.descripcionPoblacion != null)
                            //        xml.descripcionPoblacionCliente = valor;
                            //    else
                            //        xml.descripcionPoblacion = valor;
                            //    break;
                            //case "TipoVia":
                            //    if (xml.tipoVia != null)
                            //        xml.tipoViaCliente = valor;
                            //    else
                            //        xml.tipoVia = valor;
                            //    break;
                            //case "CodPostal":
                            //    if (xml.codPostal != null)
                            //        xml.codPostalCliente = valor;
                            //    else
                            //        xml.codPostal = valor;
                            //    break;
                            //case "Calle":
                            //    if (xml.calle != null)
                            //        xml.calleCliente = valor;
                            //    else
                            //        xml.calle = valor;
                            //    break;
                            //case "NumeroFinca":
                            //    if (xml.numeroFinca != null)
                            //        xml.numeroFincaCliente = valor;
                            //    else
                            //        xml.numeroFinca = valor;
                            //    break;
                            //case "AclaradorFinca":
                            //    xml.aclaradorFinca = valor;
                            //    break;
                            case "TipoIdentificador":
                                xml.tipoIdentificador = valor;
                                break;
                            case "Identificador":
                                xml.identificador = valor;
                                break;
                            case "TipoPersona":
                                xml.tipoPersona = valor;
                                break;
                            case "RazonSocial":
                                xml.razonSocial = valor;
                                break;
                            case "PrefijoPais":
                                xml.prefijoPais = valor;
                                break;
                            case "Numero":
                                xml.numero = valor;
                                break;
                            case "CorreoElectronico":
                                xml.correoElectronico = valor;
                                break;
                            case "IndicadorTipoDireccion":
                                xml.indicadorTipoDireccion = valor;
                                break;
                            case "IndicativoDeDireccionExterna":
                                xml.indicativoDeDireccionExterna = valor;
                                break;
                            case "Linea1DeLaDireccionExterna":
                                xml.linea1DeLaDireccionExterna = valor;
                                break;
                            case "Linea2DeLaDireccionExterna":
                                xml.linea2DeLaDireccionExterna = valor;
                                break;
                            case "Linea3DeLaDireccionExterna":
                                xml.linea3DeLaDireccionExterna = valor;
                                break;
                            case "Linea4DeLaDireccionExterna":
                                xml.linea4DeLaDireccionExterna = valor;
                                break;
                            case "Idioma":
                                xml.idioma = valor;
                                break;
                            case "Fecha":
                                xml.fechaActivacion = Convert.ToDateTime(valor);
                                break;
                            case "FechaActivacion":
                                xml.fechaActivacion = Convert.ToDateTime(valor);
                                break;
                            case "FechaPrevistaAccion":
                                xml.fechaActivacion = Convert.ToDateTime(valor);
                                break;
                            case "CodContrato":
                                xml.codContrato = valor;
                                break;
                            case "TipoAutoconsumo":
                                xml.tipoAutoconsumo = valor;
                                break;
                            case "TipoContratoATR":
                                xml.tipoContratoATR = valor;
                                break;
                            case "TarifaATR":
                                xml.tarifaATR = valor;
                                break;
                            case "PeriodicidadFacturacion":
                                xml.periodicidadFacturacion = valor;
                                break;
                            case "TipodeTelegestion":
                                xml.tipodeTelegestion = valor;
                                break;
                            case "Potencia":
                                xml.potenciaPeriodo[potencia] = Convert.ToDouble(valor);
                                break;
                            case "MarcaMedidaConPerdidas":
                                xml.marcaMedidaConPerdidas = valor;
                                break;
                            case "TensionDelSuministro":
                                xml.tensionDelSuministro = Convert.ToInt32(valor);
                                break;

                        }



                    

                    #endregion

                }

                if (xml.indicadorTipoDireccion == "S")
                {
                    xml.paisCliente = xml.pais;
                    xml.provinciaCliente = xml.provincia;
                    xml.municipioCliente = xml.municipio;
                    xml.codPostalCliente = xml.codPostal;
                    xml.calleCliente = xml.calle;
                    xml.tipoViaCliente = xml.tipoVia;
                    xml.numeroFincaCliente = xml.numeroFinca;
                }

                xml.linea1DeLaDireccionExterna = xml.calleCliente + " " + xml.numeroFincaCliente;
                //xml.linea2DeLaDireccionExterna = xml.acla
                xml.linea2DeLaDireccionExterna = xml.codPostalCliente + ", "
                    + provincias.DesProvincia(xml.codPostalCliente);
                xml.linea3DeLaDireccionExterna = "";
                xml.linea4DeLaDireccionExterna = "";




                
                return xml;
            }
            catch(Exception e)
            {
                              
                return null;
            }


        }
        

        //private EndesaEntity.contratacion.xxi.XML_Datos TrataXMLCompleto(string fileName)
        //{
        //    string cod_ini = "";
        //    string cod_fin = "";
        //    string valor = "";
        //    bool dentroCliente = false;

        //    int potencia = 0;

        //    EndesaEntity.contratacion.xxi.XML_Datos xml = new EndesaEntity.contratacion.xxi.XML_Datos();
        //    FileInfo file = new FileInfo(fileName);

        //    xml.fichero = file.Name;

        //    XmlTextReader r = new XmlTextReader(fileName);
        //    while (r.Read())
        //    {

        //        switch (r.NodeType)
        //        {

        //            case XmlNodeType.Element: // The node is an element.
        //                cod_ini = r.Name;

        //                if (!dentroCliente)
        //                    dentroCliente = (cod_ini == "Direccion");

        //                if (cod_ini == "Potencia")
        //                    potencia++;
        //                break;
        //            case XmlNodeType.Text: //Display the text in each element.
        //                valor = EndesaBusiness.utilidades.FuncionesTexto.ArreglaAcentos(r.Value);
        //                break;
        //            case XmlNodeType.EndElement: //Display the end of the element.
        //                cod_fin = r.Name;
        //                break;


        //        }

        //        #region XML

        //        if (cod_ini == cod_fin)
        //            switch (cod_ini)
        //            {
        //                case "NombreDePila":
        //                    xml.razonSocial = valor;
        //                    break;
        //                case "PrimerApellido":
        //                    xml.razonSocial += " " + valor;
        //                    break;
        //                case "CodigoREEEmpresaEmisora":
        //                    xml.codigoREEEmpresaEmisora = valor;
        //                    break;
        //                case "CodigoREEEmpresaDestino":
        //                    xml.codigoREEEmpresaDestino = valor;
        //                    break;
        //                case "CodigoDelProceso":
        //                    xml.codigoDelProceso = valor;
        //                    break;
        //                case "CodigoDePaso":
        //                    xml.codigoDePaso = valor;
        //                    break;
        //                case "CodigoDeSolicitud":
        //                    xml.codigoDeSolicitud = valor;
        //                    break;
        //                case "SecuencialDeSolicitud":
        //                    xml.secuencialDeSolicitud = valor;
        //                    break;
        //                case "FechaSolicitud":
        //                    xml.fechaSolicitud = Convert.ToDateTime(valor.Substring(0, 10) + " " + valor.Substring(11, 8));
        //                    break;
        //                case "CUPS":
        //                    xml.cups = valor;
        //                    break;
        //                case "CNAE":
        //                    xml.cnae = Convert.ToInt32(valor);
        //                    break;
        //                case "PotenciaExtension":
        //                    xml.potenciaExtension = Convert.ToDouble(valor);
        //                    break;
        //                case "PotenciaDeAcceso":
        //                    xml.potenciaDeAcceso = Convert.ToDouble(valor);
        //                    break;
        //                case "PotenciaInstAT":
        //                    xml.potenciaInstAT = Convert.ToDouble(valor);
        //                    break;
        //                case "IndicativoDeInterrumpibilidad":
        //                    xml.indicativoDeInterrumpibilidad = valor;
        //                    break;
        //                case "Pais":
        //                    if (xml.pais != null)
        //                        xml.paisCliente = valor;
        //                    else
        //                        xml.pais = valor;
        //                    break;
        //                case "Provincia":
        //                    if (xml.provincia != null)
        //                        xml.provinciaCliente = valor;
        //                    else
        //                        xml.provincia = valor;
        //                    break;
        //                case "Municipio":
        //                    if (xml.municipio != null)
        //                        xml.municipioCliente = valor;
        //                    else
        //                        xml.municipio = valor;
        //                    break;
        //                case "Poblacion":
        //                    if (xml.poblacion != null)
        //                        xml.poblacionCliente = valor;
        //                    else
        //                        xml.poblacion = valor;
        //                    break;
        //                case "DescripcionPoblacion":
        //                    if (xml.descripcionPoblacion != null)
        //                        xml.descripcionPoblacionCliente = valor;
        //                    else
        //                        xml.descripcionPoblacion = valor;
        //                    break;
        //                case "TipoVia":
        //                    if (xml.tipoVia != null)
        //                        xml.tipoViaCliente = valor;
        //                    else
        //                        xml.tipoVia = valor;
        //                    break;
        //                case "CodPostal":
        //                    if (xml.codPostal != null)
        //                        xml.codPostalCliente = valor;
        //                    else
        //                        xml.codPostal = valor;
        //                    break;
        //                case "Calle":
        //                    if (xml.calle != null)
        //                        xml.calleCliente = valor;
        //                    else
        //                        xml.calle = valor;
        //                    break;
        //                case "NumeroFinca":
        //                    if (xml.numeroFinca != null)
        //                        xml.numeroFincaCliente = valor;
        //                    else
        //                        xml.numeroFinca = valor;
        //                    break;
        //                case "AclaradorFinca":
        //                    xml.aclaradorFinca = valor;
        //                    break;
        //                case "TipoIdentificador":
        //                    xml.tipoIdentificador = valor;
        //                    break;
        //                case "Identificador":
        //                    xml.identificador = valor;
        //                    break;
        //                case "TipoPersona":
        //                    xml.tipoPersona = valor;
        //                    break;
        //                case "RazonSocial":
        //                    xml.razonSocial = valor;
        //                    break;
        //                case "PrefijoPais":
        //                    xml.prefijoPais = valor;
        //                    break;
        //                case "Numero":
        //                    xml.numero = valor;
        //                    break;
        //                case "CorreoElectronico":
        //                    xml.correoElectronico = valor;
        //                    break;
        //                case "IndicadorTipoDireccion":
        //                    xml.indicadorTipoDireccion = valor;
        //                    break;
        //                case "IndicativoDeDireccionExterna":
        //                    xml.indicativoDeDireccionExterna = valor;
        //                    break;
        //                case "Linea1DeLaDireccionExterna":
        //                    xml.linea1DeLaDireccionExterna = valor;
        //                    break;
        //                case "Linea2DeLaDireccionExterna":
        //                    xml.linea2DeLaDireccionExterna = valor;
        //                    break;
        //                case "Linea3DeLaDireccionExterna":
        //                    xml.linea3DeLaDireccionExterna = valor;
        //                    break;
        //                case "Linea4DeLaDireccionExterna":
        //                    xml.linea4DeLaDireccionExterna = valor;
        //                    break;
        //                case "Idioma":
        //                    xml.idioma = valor;
        //                    break;
        //                case "Fecha":
        //                    xml.fechaActivacion = Convert.ToDateTime(valor);
        //                    break;
        //                case "FechaActivacion":
        //                    xml.fechaActivacion = Convert.ToDateTime(valor);
        //                    break;
        //                case "CodContrato":
        //                    xml.codContrato = valor;
        //                    break;
        //                case "TipoAutoconsumo":
        //                    xml.tipoAutoconsumo = valor;
        //                    break;
        //                case "TipoContratoATR":
        //                    xml.tipoContratoATR = valor;
        //                    break;
        //                case "TarifaATR":
        //                    xml.tarifaATR = valor;
        //                    break;
        //                case "PeriodicidadFacturacion":
        //                    xml.periodicidadFacturacion = valor;
        //                    break;
        //                case "TipodeTelegestion":
        //                    xml.tipodeTelegestion = valor;
        //                    break;
        //                case "Potencia":
        //                    xml.potenciaPeriodo[potencia] = Convert.ToDouble(valor);
        //                    break;
        //                case "MarcaMedidaConPerdidas":
        //                    xml.marcaMedidaConPerdidas = valor;
        //                    break;
        //                case "TensionDelSuministro":
        //                    xml.tensionDelSuministro = Convert.ToInt32(valor);
        //                    break;

        //            }


        //        #endregion

        //    }

        //    return xml;


        //}

        #region comentario
        //private void GuardadoBBDD(List<EndesaEntity.contratacion.xxi.XML_Datos> lista)
        //{
        //    MySQLDB db;
        //    MySqlCommand command;
        //    string strSql = "";
        //    bool firstOnly = true;
        //    int num_reg = 0;

        //    foreach (EndesaEntity.contratacion.xxi.XML_Datos xml in lista)
        //    {
        //        if (firstOnly)
        //        {
        //            strSql = "replace into eexxi_solicitudes_tmp (CodigoREEEmpresaEmisora, CodigoREEEmpresaDestino, CodigoDelProceso,"
        //                + " CodigoDePaso, CodigoDeSolicitud, SecuencialDeSolicitud, FechaSolicitud, CUPS, CNAE, PotenciaExtension,"
        //                + " PotenciaDeAcceso, PotenciaInstAT, IndicativoDeInterrumpibilidad, Pais, Provincia, Municipio, Poblacion,"
        //                + " DescripcionPoblacion, TipoVia, CodPostal, Calle, NumeroFinca, AclaradorFinca, TipoIdentificador, Identificador,"
        //                + " TipoPersona, RazonSocial, PrefijoPais, Numero, CorreoElectronico, IndicadorTipoDireccion, IndicativoDeDireccionExterna,"
        //                + " Linea1DeLaDireccionExterna, Linea2DeLaDireccionExterna, Linea3DeLaDireccionExterna, Linea4DeLaDireccionExterna,"
        //                + " PaisCliente, ProvinciaCliente, MunicipioCliente, PoblacionCliente, DescripcionPoblacionCliente, CodPostalCliente,"
        //                + " CalleCliente, NumeroFincaCliente, PisoCliente,"
        //                + " Idioma, FechaActivacion, CodContrato, TipoAutoconsumo, TipoContratoATR, TarifaATR, PeriodicidadFacturacion, TipodeTelegestion,"
        //                + " PotenciaPeriodo1, PotenciaPeriodo2, PotenciaPeriodo3, PotenciaPeriodo4, PotenciaPeriodo5, PotenciaPeriodo6, "
        //                + " MarcaMedidaConPerdidas, TensionDelSuministro,"
        //                + " created_by, fichero) values ";
        //            firstOnly = false;
        //        }

        //        num_reg++;                          

        //        #region Campos

        //        if (xml.codigoREEEmpresaEmisora != null)
        //            strSql += "('" + xml.codigoREEEmpresaEmisora + "'";
        //        else
        //            strSql += "(null";

        //        if (xml.codigoREEEmpresaDestino != null)
        //            strSql += ", '" + xml.codigoREEEmpresaDestino + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.codigoDelProceso != null)
        //            strSql += ", '" + xml.codigoDelProceso + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.codigoDePaso != null)
        //            strSql += ", '" + xml.codigoDePaso + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.codigoDeSolicitud != null)
        //            strSql += ", '" + xml.codigoDeSolicitud + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.secuencialDeSolicitud != null)
        //            strSql += ", '" + xml.secuencialDeSolicitud + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.fechaSolicitud != null)
        //            strSql += ", '" + xml.fechaSolicitud.ToString("yyyy-MM-dd HH:mm:ss") + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cups != null)
        //            strSql += ", '" + xml.cups + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cnae != 0)
        //            strSql += ", " + xml.cnae;
        //        else
        //            strSql += ", null";


        //        if (xml.potenciaExtension != 0)
        //            strSql += ", " + xml.potenciaExtension;
        //        else
        //            strSql += ", null";

        //        if (xml.potenciaDeAcceso != 0)
        //            strSql += ", " + xml.potenciaDeAcceso;
        //        else
        //            strSql += ", null";


        //        if (xml.potenciaInstAT != 0)
        //            strSql += ", " + xml.potenciaInstAT;
        //        else
        //            strSql += ", null";

        //        if (xml.indicativoDeInterrumpibilidad != null)
        //            strSql += ", '" + xml.indicativoDeInterrumpibilidad + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.pais != null)
        //            strSql += ", '" + xml.pais + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.provincia != null)
        //            strSql += ", '" + xml.provincia + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.municipio != null)
        //            strSql += ", '" + xml.municipio + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.poblacion != null)
        //            strSql += ", '" + xml.poblacion + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.descripcionPoblacion != null)
        //            strSql += ", '" + xml.descripcionPoblacion + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.tipoVia != null)
        //            strSql += ", '" + xml.tipoVia + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.codPostal != null)
        //            strSql += ", '" + xml.codPostal + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.calle != null)
        //            strSql += ", '" + xml.calle + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.numeroFinca != null)
        //            strSql += ", '" + xml.numeroFinca + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.aclaradorFinca != null)
        //            strSql += ", '" + xml.aclaradorFinca + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.tipoIdentificador != null)
        //            strSql += ", '" + xml.tipoIdentificador + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.identificador != null)
        //            strSql += ", '" + xml.identificador + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.tipoPersona != null)
        //            strSql += ", '" + xml.tipoPersona + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.razonSocial != null)
        //            strSql += ", '" + xml.razonSocial + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.prefijoPais != null)
        //            strSql += ", '" + xml.prefijoPais + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.numero != null)
        //            strSql += ", '" + xml.numero + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.correoElectronico != null)
        //            strSql += ", '" + xml.correoElectronico + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.indicadorTipoDireccion != null)
        //            strSql += ", '" + xml.indicadorTipoDireccion + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.indicativoDeDireccionExterna != null)
        //            strSql += ", '" + xml.indicativoDeDireccionExterna + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.linea1DeLaDireccionExterna != null)
        //            strSql += ", '" + xml.linea1DeLaDireccionExterna + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.linea2DeLaDireccionExterna != null)
        //            strSql += ", '" + xml.linea2DeLaDireccionExterna + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.linea3DeLaDireccionExterna != null)
        //            strSql += ", '" + xml.linea3DeLaDireccionExterna + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.linea4DeLaDireccionExterna != null)
        //            strSql += ", '" + xml.linea4DeLaDireccionExterna + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.paisCliente != null)
        //            strSql += ", '" + xml.paisCliente + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.provinciaCliente != null)
        //            strSql += ", '" + xml.provinciaCliente + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.municipioCliente != null)
        //            strSql += ", '" + xml.municipioCliente + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.poblacionCliente != null)
        //            strSql += ", '" + xml.poblacionCliente + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.descripcionPoblacionCliente != null)
        //            strSql += ", '" + xml.descripcionPoblacionCliente + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.codPostalCliente != null)
        //            strSql += ", '" + xml.codPostalCliente + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.calleCliente != null)
        //            strSql += ", '" + xml.calleCliente + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.numeroFincaCliente != null)
        //            strSql += ", '" + xml.numeroFincaCliente + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.pisoCliente != null)
        //            strSql += ", '" + xml.pisoCliente + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.idioma != null)
        //            strSql += ", '" + xml.idioma + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.fechaActivacion != null)
        //            strSql += ", '" + xml.fechaActivacion.ToString("yyyy-MM-dd") + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.codContrato != null)
        //            strSql += ", '" + xml.codContrato + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.tipoAutoconsumo != null)
        //            strSql += ", '" + xml.tipoAutoconsumo + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.tipoContratoATR != null)
        //            strSql += ", '" + xml.tipoContratoATR + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.tarifaATR != null)
        //            strSql += ", '" + xml.tarifaATR + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.periodicidadFacturacion != null)
        //            strSql += ", '" + xml.periodicidadFacturacion + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.tipodeTelegestion != null)
        //            strSql += ", '" + xml.tipodeTelegestion + "'";
        //        else
        //            strSql += ", null";

        //        for (int i = 1; i < xml.potenciaPeriodo.Count(); i++)
        //            if (xml.potenciaPeriodo[i] != 0)
        //                strSql += ", " + xml.potenciaPeriodo[i];
        //            else
        //                strSql += ", null";

        //        if (xml.marcaMedidaConPerdidas != null)
        //            strSql += ", '" + xml.marcaMedidaConPerdidas + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.tensionDelSuministro != 0)
        //            strSql += ", " + xml.tensionDelSuministro;
        //        else
        //            strSql += ", null";

        //        strSql += ", '" + System.Environment.UserName + "'" + ", '" + xml.fichero + "'),";
        //        #endregion

        //        if(num_reg > 250)
        //        {
        //            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
        //            command = new MySqlCommand(strSql.Substring(0,strSql.Length-1), db.con);
        //            command.ExecuteNonQuery();
        //            db.CloseConnection();
        //            num_reg = 0;
        //            strSql = "";
        //            firstOnly = true;
        //        }

        //    }

        //    if (num_reg > 0)
        //    {
        //        db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
        //        command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
        //        command.ExecuteNonQuery();
        //        db.CloseConnection();
        //        num_reg = 0;
        //        strSql = "";
        //    }


        //}
        #endregion
        private void btn_word_Click(object sender, EventArgs e)
        {

            DialogResult result = MessageBox.Show("Desea generar los contratos?" + System.Environment.NewLine
                + System.Environment.NewLine,"", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {

                EndesaBusiness.contratacion.eexxi.EEXXI xxi = new EndesaBusiness.contratacion.eexxi.EEXXI();
                Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic = new Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos>();
                // dic = xxi.Altas();

                foreach (KeyValuePair<string, EndesaEntity.contratacion.xxi.XML_Datos> p in dic)
                {

                }

                MessageBox.Show("Fin.",
                    "Exportacion fin",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }

        }

        private void FrmEEXXI_Load(object sender, EventArgs e)
        {

            //facturas = new EndesaBusiness.contratacion.eexxi.Facturacion();
            //xxi = new EndesaBusiness.contratacion.eexxi.EEXXI();
            //// Carga los distintos estados de los casos que podemos seleccionar en el formulario
            //estados_casos = new EndesaBusiness.contratacion.eexxi.EstadosCasos();

            //DateTime fecha = EndesaBusiness.utilidades.FechasCargas.UltimaActualizacionPS("PSAT");
            //DateTime ultima_solicitud = EndesaBusiness.utilidades.FechasCargas.UltimaSolicitud_EEXXI_XML();
            //this.lbl_ultima_solicitud.Text = string.Format("Última solicitud: {0}", ultima_solicitud);
            //this.lblPSAT.Text = string.Format("Última actualización de PS_AT: {0}", fecha);

            //if (fecha.Date == DateTime.Now.Date)
            //    this.lblPSAT.ForeColor = Color.Green;
            //else
            //    this.lblPSAT.ForeColor = Color.Red;

            provincias = new EndesaBusiness.global.Provincias("eexxi_param_provincias", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
            municipio = new EndesaBusiness.global.Municipios();

            Cursor.Current = Cursors.WaitCursor;
            LoadData();
            Cursor.Current = Cursors.Default;
        }

        private void LoadData()
        {

            // xxi = new GO.contratacion.eexxi.EEXXI();
            // xxi.PuntosActivos();

            
            xxi = new EndesaBusiness.contratacion.eexxi.EEXXI();
            // Carga los distintos estados de los casos que podemos seleccionar en el formulario
            estados_casos = new EndesaBusiness.contratacion.eexxi.EstadosCasos();

            DateTime fecha = EndesaBusiness.utilidades.FechasCargas.UltimaActualizacionPS("PSAT");
            DateTime ultima_solicitud = EndesaBusiness.utilidades.FechasCargas.UltimaSolicitud_EEXXI_XML();
            this.lbl_ultima_solicitud.Text = string.Format("Última solicitud: {0}", ultima_solicitud);
            this.lblPSAT.Text = string.Format("Última actualización de PS_AT: {0}", fecha);

            inventario = new EndesaBusiness.contratacion.eexxi.Inventario();

            casos = new EndesaBusiness.contratacion.eexxi.Casos(inventario);
            //casos.CargaCasosAbiertos();


            // dgv
            dgv.AutoGenerateColumns = false;
            DataGridViewComboBoxColumn cboBoxColumn = (DataGridViewComboBoxColumn)dgv.Columns[11];
            cboBoxColumn.DataSource = this.estados_casos.dic.Values.ToList();
            cboBoxColumn.ValueMember = "estado_id";
            cboBoxColumn.DisplayMember = "descripcion";
            cboBoxColumn.DataPropertyName = "estado_id";
            dgv.DataSource = casos.dic.Values.ToList().OrderBy(z => z.nif).ToList();

            // Altas para facturar
            LoadDataFacturas();

            //lbl_dgv_facturar_total.Text = string.Format("Total Registros: {0:#,##0}", dgv_Inventario.RowCount);

            dgv_Inventario.AutoGenerateColumns = false;
            dgv_Inventario.DataSource = inventario.dic.Values.Where(z => z.vigente).ToList();
            lbl_total_contratos.Text = string.Format("Total Registros: {0:#,##0}", dgv_Inventario.RowCount);
            
            lbl_total_resumen.Text = string.Format("Total Registros: {0:#,##0}", casos.dic.Values.ToList().Count());
        }
              
        private void LoadDataFacturas()
        {
            facturas = null;
            facturas = new EndesaBusiness.contratacion.eexxi.Facturacion();
            dgv_facturar.AutoGenerateColumns = false;
            //dgv_facturar.DataSource = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.fecha_envio_mail).ToList();
            dgv_facturar.DataSource = facturas.dic_informe.Values.ToList().OrderBy(z => z.fecha_envio_mail).ToList();
            lbl_total_altas_facturar.Text = string.Format("Total Registros: {0:#,##0}", facturas.dic_informe.Values.ToList().Count());
        }

        private void cargaXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CargaXML();                       
        }

        private void pruebaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            xxi.ProcesaSolicitudes();
            //xxi.GeneraExcelComprobacion();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void herramientasToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void generarContratoYCartaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> dic_imprime_contratos = new Dictionary<string, string>();
            
            foreach (DataGridViewRow row in dgv.SelectedRows)
            {
                DataGridViewCell cups = (DataGridViewCell)
                 row.Cells[2];
                DataGridViewCell codigo_solicitud = (DataGridViewCell)
                 row.Cells[3];
                // Buscamos el XML de alta para poner imprimir los documentos
                EndesaEntity.contratacion.xxi.Casos_Tabla_Historico o;
                if (casos.dic_casos_altas.TryGetValue(cups.Value.ToString(), out o))
                    dic_imprime_contratos.Add(cups.Value.ToString(), o.codigo_solicitud);
                
            }           

            
            if(dic_imprime_contratos.Count > 0)
            {
                DialogResult result2 = MessageBox.Show("¿Desea generar el contrato y Carta en PDF para los registros seleccionados?", "Generar documentación",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    xxi.ImprimeContratos(dic_imprime_contratos);

                    GuardaCambios_dgv();
                    LoadData();

                    MessageBox.Show("Generación de contratos finalizada.",
                             "Contratos",
                             MessageBoxButtons.OK,
                             MessageBoxIcon.Information);
                }
                    
            }
            
        }

        private void dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            
            if (e.RowIndex >= 0)
            {
                cups20 = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[2];
                codigo_solicitud = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[3];               

            }
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {

            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.FileName = "Casos_EEXXI_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                casos.CasosAbiertos_a_Excel(save.FileName);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    // System.Diagnostics.Process.Start(save.FileName);
                }
            }

        }

        private void dgv_CellEnter(object sender, DataGridViewCellEventArgs e)
        {

            if (e.ColumnIndex == 11)
            {
                this.dgv.Rows[e.RowIndex].Cells[13].Value = System.Environment.UserName.ToUpper();
                btnSave.Enabled = true;
            }
        }
        
        private void GuardaCambios_dgv()
        {

            bool guardarCambios = false;

            foreach (DataGridViewRow row in dgv.Rows)
            {

                if (row.Cells[13].Value != null)
                    guardarCambios = true;
            }

            if (guardarCambios)
            {

                DialogResult result_1 = MessageBox.Show("¿Desea guardar los cambios?",                
                "Cambios en casos pendientes EEXXI", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result_1 == DialogResult.Yes)
                    foreach (DataGridViewRow row in dgv.Rows)
                    {
                        string cups = "";
                        string codigo_solicitud = "";
                        int estado_id;
                        if (row.Cells[13].Value != null)
                        {
                            if (row.Cells[11].Value != null)
                            {
                                cups = row.Cells[2].Value.ToString();
                                codigo_solicitud = row.Cells[3].Value.ToString();
                                estado_id = Convert.ToInt32(row.Cells[12].Value);
                                casos.ActualizaEstadoCaso(cups, codigo_solicitud, estado_id);
                                
                            }
                        }
                        row.Cells[13].Value = null;
                    }



            }

            
        }
        private void FrmEEXXI_FormClosing(object sender, FormClosingEventArgs e)
        {
            GuardaCambios_dgv();
            usage.End("Contratación", "FrmEndesaEnergia21" ,"N/A");
        }

        private void enviarInformeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.contratacion.eexxi.Informes informe = new EndesaBusiness.contratacion.eexxi.Informes(inventario);
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {

        }

        private void Dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            GuardaCambios_dgv();
            btnSave.Enabled = false;
        }

        

        private void ModificarDireccionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> dic_imprime_contratos = new Dictionary<string, string>();
            forms.contratacion.FrmDirecciones ff = new FrmDirecciones();
            foreach (DataGridViewRow row in dgv.SelectedRows)
            {
                DataGridViewCell cups = (DataGridViewCell)
                 row.Cells[2];
                DataGridViewCell codigo_solicitud = (DataGridViewCell)
                 row.Cells[3];
                // Buscamos el XML de alta para poner imprimir los documentos
                EndesaEntity.contratacion.xxi.Casos_Tabla_Historico o;
                if (casos.dic_casos_altas.TryGetValue(cups.Value.ToString(), out o))
                    dic_imprime_contratos.Add(cups.Value.ToString(), o.codigo_solicitud);

            }

            ff.dic = dic_imprime_contratos;
            ff.xxi = xxi;
            ff.ShowDialog();

        }

        private void Dgv_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
           
        }

        private void Dgv_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            forms.contratacion.FrmComentarios f = new FrmComentarios();            
            f.id_codigo = dgv.CurrentRow.Cells[2].Value.ToString() + dgv.CurrentRow.Cells[3].Value.ToString();
            f.ShowDialog();
            LoadData();
        }

        private void btn_inventario_excel_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                inventario.Inventario_a_Excel(save.FileName);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }
        }

        private void parámetrosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización Endesa Energía XXI";
            p.tabla = "eexxi_param";
            p.esquemaString = "CON";
            p.Show();
        }

        private void generaXMLT105AntiguoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.xml.XMLFunciones xml_t105 = new EndesaBusiness.xml.XMLFunciones();
            //xml_t105.CargaXML();
        }

        private void generarContratoYCartaAgrupadoNIFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Dictionary<string, List<EndesaEntity.contratacion.xxi.Cups_Solicitud>> dic_imprime_contratos =
                new Dictionary<string, List<EndesaEntity.contratacion.xxi.Cups_Solicitud>>();

            EndesaBusiness.contratacion.eexxi.Solicitudes solicitudes = 
                new EndesaBusiness.contratacion.eexxi.Solicitudes("T1", "05");           


            foreach (DataGridViewRow row in dgv.SelectedRows)
            {
                DataGridViewCell nif = (DataGridViewCell)
                 row.Cells[0];
                DataGridViewCell cups = (DataGridViewCell)
                 row.Cells[2];
                DataGridViewCell codigo_solicitud = (DataGridViewCell)
                 row.Cells[3];

                // Buscamos el XML de alta para poner imprimir los documentos
                EndesaEntity.contratacion.xxi.Casos_Tabla_Historico o;
                if (casos.dic_casos_altas.TryGetValue(cups.Value.ToString(), out o))
                {
                    List<EndesaEntity.contratacion.xxi.Cups_Solicitud> oo;
                    if (!dic_imprime_contratos.TryGetValue(nif.Value.ToString(), out oo))
                    {
                        oo = new List<EndesaEntity.contratacion.xxi.Cups_Solicitud>();
                        EndesaEntity.contratacion.xxi.Cups_Solicitud c = new EndesaEntity.contratacion.xxi.Cups_Solicitud();
                        c.cups = cups.Value.ToString();
                        //c.solicitud = codigo_solicitud.Value.ToString();
                        solicitudes.GetSolicitud(c.cups);
                        if (solicitudes.existe)
                            c.solicitud = solicitudes.codigoDeSolicitud;
                        else
                            c.solicitud = "";
                        oo.Add(c);
                        dic_imprime_contratos.Add(nif.Value.ToString(), oo);
                    }
                    else
                    {
                        EndesaEntity.contratacion.xxi.Cups_Solicitud c = new EndesaEntity.contratacion.xxi.Cups_Solicitud();
                        c.cups = cups.Value.ToString();
                        //c.solicitud = codigo_solicitud.Value.ToString();
                        solicitudes.GetSolicitud(c.cups);
                        if (solicitudes.existe)
                            c.solicitud = solicitudes.codigoDeSolicitud;
                        else
                            c.solicitud = "";
                        oo.Add(c);
                    }
                        
                }
                    

            }

            if (dic_imprime_contratos.Count > 0)
            {
                DialogResult result2 = MessageBox.Show("¿Desea generar el contrato y Carta en PDF para los registros seleccionados?", "Generar documentación",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    xxi.ImprimeContratosAgrupadosNIF(dic_imprime_contratos);

                    GuardaCambios_dgv();
                    LoadData();

                    MessageBox.Show("Generación de contratos finalizada.",
                             "Contratos",
                             MessageBoxButtons.OK,
                             MessageBoxIcon.Information);
                }

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int filas_seleccionadas = 0;

            Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion> d =
                new Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion>();

            foreach (DataGridViewRow row in dgv_facturar.SelectedRows)
            {
                filas_seleccionadas++;
                DataGridViewCell cups = (DataGridViewCell)
                 row.Cells[0];
                DataGridViewCell fecha_alta = (DataGridViewCell)
                 row.Cells[3];

                EndesaEntity.contratacion.xxi.Informe_Facturacion o;
                if (facturas.dic_informe.TryGetValue(cups.Value.ToString(), out o))
                    d.Add(cups.Value.ToString(), o);

            }

            if(filas_seleccionadas > 0)
            {
                Cursor.Current = Cursors.WaitCursor;
                string rutaSalida = @"c:\temp\ALTAS_EEXXI_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
                facturas.InformeViaMail(rutaSalida, true, d);
                Cursor.Current = Cursors.Default;
                LoadDataFacturas();
            }
            else
            {
                MessageBox.Show("No ha seleccionado ninguna fila para incluir en el informe."
                    + System.Environment.NewLine
                    + "Seleccione algún registro para mostrarlo en el informe."
                    , "Altas para facurar",
                  MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

           

        }

        private void btn_facturas_altas_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.FileName = "Altas_para_facturar_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                facturas.InformeViaMail(save.FileName, false, facturas.dic_informe); 
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                     System.Diagnostics.Process.Start(save.FileName);
                }
            }
        }

        private void dgv_facturar_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dgv_facturar.Columns[e.ColumnIndex].SortMode != DataGridViewColumnSortMode.NotSortable)
            {
                if (e.ColumnIndex == newSortColumn)
                {
                    if (newColumnDirection == ListSortDirection.Ascending)
                        newColumnDirection = ListSortDirection.Descending;
                    else
                        newColumnDirection = ListSortDirection.Ascending;
                }

                newSortColumn = e.ColumnIndex;
                                
                dgv_facturar.AutoGenerateColumns = false;
                dgv_facturar.DataSource = OrdenaColumna(dgv_facturar.Columns[e.ColumnIndex].DataPropertyName, newColumnDirection); ;
                

            }
        }

        private List<EndesaEntity.contratacion.xxi.Informe_Facturacion>
            OrdenaColumna(string columna, ListSortDirection direccion)
        {
            List<EndesaEntity.contratacion.xxi.Informe_Facturacion> l =
                new List<EndesaEntity.contratacion.xxi.Informe_Facturacion>();

            switch (columna)
            {
                case "cups":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.cups).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.cups).ToList();
                    break;

                case "nif":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.nif).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.nif).ToList();
                    break;

                case "cliente":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.cliente).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.cliente).ToList();
                    break;
                case "fecha_alta":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.fecha_alta).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.fecha_alta).ToList();
                    break;

                case "fecha_baja":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.fecha_baja).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.fecha_baja).ToList();
                    break;

                case "p1":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.p1).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.p1).ToList();
                    break;
                case "p2":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.p2).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.p2).ToList();
                    break;
                case "p3":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.p3).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.p3).ToList();
                    break;
                case "p4":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.p4).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.p4).ToList();
                    break;
                case "p5":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.p5).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.p5).ToList();
                    break;
                case "p6":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.p6).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.p6).ToList();
                    break;

                case "tarifa":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.tarifa).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.tarifa).ToList();
                    break;

                case "tension":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.tension).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.tension).ToList();
                    break;
                case "fecha_envio_mail":
                    if (direccion == ListSortDirection.Ascending)
                        l = facturas.dic_informe.Values.ToList().OrderBy(z => z.fecha_envio_mail).ToList();
                    else
                        l = facturas.dic_informe.Values.ToList().OrderByDescending(z => z.fecha_envio_mail).ToList();
                    break;


            }

            return l;

        }

        private void enviarInformePorEmailToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion> d =
               new Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion>();

            foreach (DataGridViewRow row in dgv_facturar.SelectedRows)
            {
                DataGridViewCell cups = (DataGridViewCell)
                 row.Cells[0];
                DataGridViewCell fecha_alta = (DataGridViewCell)
                 row.Cells[3];

                EndesaEntity.contratacion.xxi.Informe_Facturacion o;
                if (facturas.dic_informe.TryGetValue(cups.Value.ToString(), out o))
                    d.Add(cups.Value.ToString(), o);

            }

            Cursor.Current = Cursors.WaitCursor;
            string rutaSalida = @"c:\temp\ALTAS_EEXXI_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
            facturas.InformeViaMail(rutaSalida, true, d);
            Cursor.Current = Cursors.Default;
            LoadData();
        }

        private void dgv_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dgv.Columns[e.ColumnIndex].SortMode != DataGridViewColumnSortMode.NotSortable)
            {
                if (e.ColumnIndex == newSortColumn)
                {
                    if (newColumnDirection == ListSortDirection.Ascending)
                        newColumnDirection = ListSortDirection.Descending;
                    else
                        newColumnDirection = ListSortDirection.Ascending;
                }

                newSortColumn = e.ColumnIndex;

                dgv.AutoGenerateColumns = false;
                dgv.DataSource = OrdenaColumnaPrincipal(dgv.Columns[e.ColumnIndex].DataPropertyName, newColumnDirection); 


            }
        }

        private List<EndesaEntity.contratacion.xxi.Casos_Tabla_Historico>
            OrdenaColumnaPrincipal(string columna, ListSortDirection direccion)
        {
            List<EndesaEntity.contratacion.xxi.Casos_Tabla_Historico> l =
                new List<EndesaEntity.contratacion.xxi.Casos_Tabla_Historico>();

            switch (columna)
            {

                case "nif":
                    if (direccion == ListSortDirection.Ascending)
                        l = casos.dic.Values.ToList().OrderBy(z => z.nif).ToList();
                    else
                        l = casos.dic.Values.ToList().OrderByDescending(z => z.nif).ToList();
                    break;

                case "razon_social":
                    if (direccion == ListSortDirection.Ascending)
                        l = casos.dic.Values.ToList().OrderBy(z => z.razon_social).ToList();
                    else
                        l = casos.dic.Values.ToList().OrderByDescending(z => z.razon_social).ToList();
                    break;

                case "cups":
                    if (direccion == ListSortDirection.Ascending)
                        l = casos.dic.Values.ToList().OrderBy(z => z.cups).ToList();
                    else
                        l = casos.dic.Values.ToList().OrderByDescending(z => z.cups).ToList();
                    break;        
                
                case "fecha_alta":
                    if (direccion == ListSortDirection.Ascending)
                        l = casos.dic.Values.ToList().OrderBy(z => z.fecha_alta).ToList();
                    else
                        l = casos.dic.Values.ToList().OrderByDescending(z => z.fecha_alta).ToList();
                    break;

                case "fecha_baja":
                    if (direccion == ListSortDirection.Ascending)
                        l = casos.dic.Values.ToList().OrderBy(z => z.fecha_baja).ToList();
                    else
                        l = casos.dic.Values.ToList().OrderByDescending(z => z.fecha_baja).ToList();
                    break;

                case "existe_alta":
                    if (direccion == ListSortDirection.Ascending)
                        l = casos.dic.Values.ToList().OrderBy(z => z.existe_alta).ToList();
                    else
                        l = casos.dic.Values.ToList().OrderByDescending(z => z.existe_alta).ToList();
                    break;

                case "empresa_ps":
                    if (direccion == ListSortDirection.Ascending)
                        l = casos.dic.Values.ToList().OrderBy(z => z.empresa_ps).ToList();
                    else
                        l = casos.dic.Values.ToList().OrderByDescending(z => z.empresa_ps).ToList();
                    break;

                case "estado_contrato_ps":
                    if (direccion == ListSortDirection.Ascending)
                        l = casos.dic.Values.ToList().OrderBy(z => z.estado_contrato_ps).ToList();
                    else
                        l = casos.dic.Values.ToList().OrderByDescending(z => z.estado_contrato_ps).ToList();
                    break;

                case "crear_incidencia":
                    if (direccion == ListSortDirection.Ascending)
                        l = casos.dic.Values.ToList().OrderBy(z => z.crear_incidencia).ToList();
                    else
                        l = casos.dic.Values.ToList().OrderByDescending(z => z.crear_incidencia).ToList();
                    break;

                case "documentacion_impresa":
                    if (direccion == ListSortDirection.Ascending)
                        l = casos.dic.Values.ToList().OrderBy(z => z.documentacion_impresa).ToList();
                    else
                        l = casos.dic.Values.ToList().OrderByDescending(z => z.documentacion_impresa).ToList();
                    break;

                case "estado_id":
                    if (direccion == ListSortDirection.Ascending)
                        l = casos.dic.Values.ToList().OrderBy(z => z.estado_id).ToList();
                    else
                        l = casos.dic.Values.ToList().OrderByDescending(z => z.estado_id).ToList();
                    break;

                case "comentario":
                    if (direccion == ListSortDirection.Ascending)
                        l = casos.dic.Values.ToList().OrderBy(z => z.comentario).ToList();
                    else
                        l = casos.dic.Values.ToList().OrderByDescending(z => z.comentario).ToList();
                    break;                


            }

            return l;

        }

        private void importaciónExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = DialogResult.Yes;
            EndesaBusiness.contratacion.eexxi.Aviso_a_COR cor
                = new EndesaBusiness.contratacion.eexxi.Aviso_a_COR();

            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Seleccione archivos Excel de respuesta COR";
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                result = MessageBox.Show("¿Desea procesar los Excels seleccionados?",
                  "Importación ficheros Excel repuesta COR",
                  MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    foreach (string fileName in d.FileNames)
                    {
                        cor.CargaExcelRespuestaCOR(fileName);
                    }                    
                    Cursor.Current = Cursors.Default;

                    MessageBox.Show("Proceso Finalizado."
                       + System.Environment.NewLine                       
                       + System.Environment.NewLine
                       + p.GetValue("salida_excels", DateTime.Now, DateTime.Now),
                 "Importación ficheros Excel respuesta COR",
                 MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void generaXMLT105AntiguoAPartirDeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.xml.XMLFunciones xml_t105 = new EndesaBusiness.xml.XMLFunciones();
            xml_t105.CargaXML_A302_A305();
        }

        private void cargaExcelEventualesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = DialogResult.Yes;

            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Seleccione archivos Excel de Eventuales";
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

            }
        }

        private void generaExcelPasoACORToolStripMenuItem_Click(object sender, EventArgs e)
        {

            FrmInformePaso_a_COR f = new FrmInformePaso_a_COR();
            f.Show();

          
        }

        private void cargaXMLExcepcionalSinRestriccionesDeTarifaNiPotenciaP6ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CargaXML_SinRestricciones();
        }

        private void generaExcelAvisoVentasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmInformeAviso_Ventas f = new FrmInformeAviso_Ventas();
            f.Show();
        }

        private void avisoVentasManualToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmEndesaEnergia21_AvisoVentas_Manual f = new FrmEndesaEnergia21_AvisoVentas_Manual();
            f.Show();
        }

        private void eliminarAltasFacturartoolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaEntity.contratacion.xxi.Informe_Facturacion fact = new EndesaEntity.contratacion.xxi.Informe_Facturacion();
            int numero_registros_seleccionados = dgv_facturar.SelectedRows.Count;


            if (numero_registros_seleccionados <= 0)
            {
                MessageBox.Show("No ha seleccionado ningún registro", "Eliminar alta para facturar", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {

                DialogResult result = MessageBox.Show("¿Está seguro que quiere eliminar " + (numero_registros_seleccionados > 1 ? "los " + numero_registros_seleccionados + " registros seleccionados" : "el registro seleccionado") + "?", "Eliminar registros",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if ( (result == DialogResult.Yes))
                {
                    Cursor.Current = Cursors.WaitCursor;

                    foreach (DataGridViewRow row in dgv_facturar.SelectedRows)
                    {
                        DataGridViewCell cups = (DataGridViewCell)
                         row.Cells[0];
                        DataGridViewCell nif = (DataGridViewCell)
                         row.Cells[1];
                        DataGridViewCell cliente = (DataGridViewCell)
                         row.Cells[2];
                        DataGridViewCell fecha_alta = (DataGridViewCell)
                         row.Cells[3];
                        DataGridViewCell fecha_baja = (DataGridViewCell)
                         row.Cells[4];

                        //MessageBox.Show("Eliminando... CUP: " + cups.Value.ToString() + " | FECHA ALTA: " + ((DateTime) fecha_alta.Value).ToShortDateString() + " | FECHA BAJA: " + ((DateTime)fecha_baja.Value).ToShortDateString(), "Eliminar alta para facturar",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    
                        fact.cups = cups.Value.ToString();
                        fact.nif = nif.Value.ToString();
                        fact.cliente = cliente.Value.ToString();
                        fact.fecha_alta = (DateTime)fecha_alta.Value;
                        fact.fecha_baja = (DateTime)fecha_baja.Value;

                        facturas.AddNoFacturar(fact);

                        facturas.dic_informe.Remove(cups.Value.ToString());
                       

                    }

                    dgv_facturar.DataSource = facturas.dic_informe.Values.ToList().OrderBy(z => z.fecha_envio_mail).ToList();
                    lbl_total_altas_facturar.Text = string.Format("Total Registros: {0:#,##0}", facturas.dic_informe.Values.ToList().Count());

                    dgv_facturar.Update();
                    dgv_facturar.Refresh();

                    Cursor.Current = Cursors.Default;

                    MessageBox.Show( (numero_registros_seleccionados > 1 ? numero_registros_seleccionados + " registros eliminados" : "Registro eliminado"), "Eliminar registros",
                       MessageBoxButtons.OK, MessageBoxIcon.Information);

                   
                }
                
            }


        }
    }
}



