
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.medida
{
    public partial class FrmAdif_Importar : Form
    {
        EndesaBusiness.adif.ProcesosFunciones pf = new EndesaBusiness.adif.ProcesosFunciones();
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmAdif_Importar()
        {
            usage.Start("Medida", "FrmAdif_Importar", "N/A");
            InitializeComponent();
            ActualizaFechas();
        }

        private void ActualizaFechas()
        {
            DateTime adif = new DateTime();
            DateTime curvas = new DateTime();
            DateTime resumen = new DateTime();            

            adif = pf.GetLastProcess("Importar PO1011");
            curvas = pf.GetLastProcess("Importar CC");
            resumen = pf.GetLastProcess("Importar CR");

            lblAdif.Text = string.Format("{0}", adif.ToString("dd/MM/yyyy"));
            lblAdif.ForeColor = GetColor(adif);
            lblCuartoHoraria.Text = string.Format("{0}", curvas.ToString("dd/MM/yyyy"));
            lblCuartoHoraria.ForeColor = GetColor(curvas);
            lblResumen.Text = string.Format("{0}", resumen.ToString("dd/MM/yyyy"));
            lblResumen.ForeColor = GetColor(resumen);

        }

        private Color GetColor(DateTime d)
        {
           int diff;

           diff = (DateTime.Now - d).Days;

            if (diff <= 1)
                return Color.Green;
            else if (diff > 1 && diff < 4)
                return Color.Orange;
            else
                return Color.Red;

        }

        private void btnFicherosDat_Click(object sender, EventArgs e)
        {
            int totalArchivos = 0;
            forms.FrmProgressBar pb = new forms.FrmProgressBar();
            int i = 0;
            DateTime begin = new DateTime();
            int total_cups = 0;

            EndesaBusiness.adif.AdifImportar adif_imp = new EndesaBusiness.adif.AdifImportar();

            begin = DateTime.Now;
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "ficheros formato PO.10.11|*.*";
            d.Multiselect = true;
            EndesaBusiness.adif.P01011_Funciones pp = new EndesaBusiness.adif.P01011_Funciones();
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                totalArchivos = d.FileNames.Count(); ;

                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = d.FileNames.Count();
                pb.Text = "Importando... ";

                foreach (string fileName in d.FileNames)
                {
                    i++;
                    if (i == 100)
                    {
                        i = 0;
                        pp.Save();
                        pp = null;
                        pp = new EndesaBusiness.adif.P01011_Funciones();                        
                    }

                    pb.txtDescripcion.Text = "Importando: " + fileName.ToString();
                    pb.progressBar.Increment(1);
                    pb.Refresh();
                                    
                    pp.CargaP01011(fileName);
                    pp.GuardaP01011(fileName);

                }
               
                pb.Close();
                pp.Save();

                total_cups = adif_imp.TratarCarga();
                pf.SaveProcess("Importar PO1011", "Importacion curvas ADIF", begin, DateTime.Now);

                MessageBox.Show("La importación ha concluido correctamente." 
                    + System.Environment.NewLine + System.Environment.NewLine 
                    + "Ficheros Seleccionados: " + totalArchivos + System.Environment.NewLine
                    + "Ficheros cargados: " + totalArchivos + System.Environment.NewLine
                    + "CUPS cargados: " + total_cups,
                    "Importación ficheros ficheros formato PO.10.11 ADIF",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        

        private void lblAdif_Click(object sender, EventArgs e)
        {
                       
        }



        private void btnCuartoHoraria_Click(object sender, EventArgs e)
        {
            EndesaBusiness.medida.CurvaCuartoHorariaFunciones cc = new EndesaBusiness.medida.CurvaCuartoHorariaFunciones();
            EndesaBusiness.adif.Param p = new EndesaBusiness.adif.Param();
            FileInfo file;
            FileInfo fileDes;
            DateTime begin = new DateTime();
            string prefijoCurvaResumen;
            string[] listaArchivos;
            EndesaBusiness.utilidades.ZipUnZip zip = new EndesaBusiness.utilidades.ZipUnZip();


            prefijoCurvaResumen = "*" + p.GetParam("PrefijoCC") + "*.TXT";
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "ficheros Curvas CuartoHoraria|*.ZIP";
            d.Multiselect = false;
            EndesaBusiness.adif.P01011_Funciones pp = new EndesaBusiness.adif.P01011_Funciones();
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                begin = DateTime.Now;
                Cursor.Current = Cursors.WaitCursor;
                foreach (string fileName in d.FileNames)
                {
                    file = new FileInfo(fileName);
                    if (file.Extension.ToUpper() == ".ZIP")
                    {
                        DirectoryInfo dir = new DirectoryInfo(p.GetParam("UnzipFolder"));
                        if (!dir.Exists)
                            dir.Create();
                        fileDes = new FileInfo(dir.FullName + file.Name);
                        if (fileDes.Exists)
                            fileDes.Delete();
                        file.CopyTo(dir.FullName + "//" + file.Name);

                        zip.DescomprimirArchivo(fileDes.FullName);
                        // Buscamos los archivos descomprimidos
                        listaArchivos = Directory.GetFiles(dir.FullName, prefijoCurvaResumen);
                        for (int j = 0; j < listaArchivos.Length; j++)
                        {
                            cc.ImportarCuartoHorariaPorLinea(listaArchivos[j]);
                        }
                    }
                    else
                        cc.ImportarCuartoHorariaPorLinea(fileName);
                }
                pf.SaveProcess("Importar CC", "Importación del archivo de CC", begin, DateTime.Now);
                Cursor.Current = Cursors.Default;
                this.ActualizaFechas();

                MessageBox.Show("Importación finalizada."                  
                  ,"Importación de Curvas CuartoHorarias",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Information);

            }
        }
    

        private void btnResumen_Click(object sender, EventArgs e)
        {

            EndesaBusiness.medida.CurvaResumenFunciones cr = new EndesaBusiness.medida.CurvaResumenFunciones();
            EndesaBusiness.adif.Param p = new EndesaBusiness.adif.Param();
            FileInfo file;
            FileInfo fileDes;
            DateTime begin = new DateTime();
            string prefijoCurvaResumen;
            string[] listaArchivos;
            EndesaBusiness.utilidades.ZipUnZip zip = new EndesaBusiness.utilidades.ZipUnZip();


            prefijoCurvaResumen = "*" + p.GetParam("PrefijoCR") + "*.TXT";
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "ficheros Curvas Resumen|*.ZIP";
            d.Multiselect = false;
            EndesaBusiness.adif.P01011_Funciones pp = new EndesaBusiness.adif.P01011_Funciones();
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                begin = DateTime.Now;
                Cursor.Current = Cursors.WaitCursor;
                foreach (string fileName in d.FileNames)
                {
                    file = new FileInfo(fileName);
                    if(file.Extension.ToUpper() == ".ZIP")
                    {
                        DirectoryInfo dir = new DirectoryInfo(p.GetParam("UnzipFolder"));
                        if (!dir.Exists)
                            dir.Create();
                        fileDes = new FileInfo(dir.FullName + file.Name);
                        if (fileDes.Exists)
                            fileDes.Delete();
                        file.CopyTo(dir.FullName + "//" + file.Name);                        
                        zip.DescomprimirArchivo(fileDes.FullName);
                        // Buscamos los archivos descomprimidos
                        listaArchivos = Directory.GetFiles(dir.FullName, prefijoCurvaResumen);
                        for (int j = 0; j < listaArchivos.Length; j++)
                        {
                            cr.CargaResumen(listaArchivos[j]);
                        }
                    }
                    else
                        cr.CargaResumen(fileName);
                }
                pf.SaveProcess("Importar CR", "Importación del archivo de CR", begin, DateTime.Now);
                Cursor.Current = Cursors.Default;
                this.ActualizaFechas();
                MessageBox.Show("Importación finalizada.",
                  "Importación de Curvas Resumen",                 
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Information);


            }
        }

        private void cerrarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void FrmAdif_Importar_Load(object sender, EventArgs e)
        {
            txtYYYYMM.Text = DateTime.Now.AddMonths(-1).ToString("yyyyMM");
        }

        private void btnDescargaFTP_Click(object sender, EventArgs e)
        {
            
            
            List<string> rutas_validas = new List<string>();
            DateTime f = new DateTime();
            bool continuar;

            string dirFTP = "";
            List<string> lista_carpetas;
            List<string> lista_archivos;

            string mes_letra = "";
            string year = "";
            string subruta = "";

            EndesaBusiness.utilidades.Fechas utilFecha = new EndesaBusiness.utilidades.Fechas();
            forms.FrmProgressBar pb = new forms.FrmProgressBar();
            double percent = 0;
            int totalArchivos = 0;
            int totalArchivosDescargados = 0;
            int jj = 0;

            EndesaBusiness.utilidades.UltimateFTP ftpClient;

            try
            {

                EndesaBusiness.utilidades.Param param = 
                    new EndesaBusiness.utilidades.Param("adif_param", EndesaBusiness.servidores.MySQLDB.Esquemas.MED);

                DialogResult result = MessageBox.Show("Aviso muy importante:" + System.Environment.NewLine
                + System.Environment.NewLine
                + "La descarga de archivos desde el FTP de ADIF puede durar varios minutos." + System.Environment.NewLine 
                + "La ruta por defecto donde se descargarán los archivos es " + param.GetValue("ftp_local_dir", DateTime.Now, DateTime.Now)
                + System.Environment.NewLine
                + System.Environment.NewLine
                + "¿Desea continuar con la descarga? " + System.Environment.NewLine
                , "Descarga archivos FTP ADIF",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                continuar = (result == DialogResult.Yes);
                if (continuar)
                {
                    
                    f = new DateTime(Convert.ToInt32(txtYYYYMM.Text.Substring(0, 4)), Convert.ToInt32(txtYYYYMM.Text.Substring(4, 2)), 1);
                    mes_letra = utilFecha.ConvierteMes_a_Letra(f);
                    year = f.Year.ToString();


                    ftpClient = new EndesaBusiness.utilidades.UltimateFTP(
                        param.GetValue("ftp_server"),
                        param.GetValue("ftp_user"),
                        param.GetValue("ftp_pass"),
                        param.GetValue("ftp_port"));


                    //EndesaBusiness.utilidades.FTP ftpClient = new EndesaBusiness.utilidades.FTP(param.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                    //    param.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                    //    param.GetValue("ftp_pass", DateTime.Now, DateTime.Now));

                    //utilidades.FTP ftpClient = new GO.utilidades.FTP(param.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                    //param.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                    //param.GetValue("ftp_pass", DateTime.Now, DateTime.Now),
                    //param.GetValue("ftp_proxy", DateTime.Now, DateTime.Now),
                    //param.GetValue("ftp_proxy_port", DateTime.Now, DateTime.Now),
                    //param.GetValue("ftp_proxy_port", DateTime.Now, DateTime.Now));

                    dirFTP = param.GetValue("ftp_main_folder", DateTime.Now, DateTime.Now)
                        + year + "/GRUPO_17_CURVAS_ADIF_MES_" + mes_letra + "_" + year;

                    //lista_carpetas = ftpClient.ListaDirectorios(dirFTP);

                    //for (int i = 0; i < lista_carpetas.Count(); i++)
                    //{
                    //    if (lista_carpetas[i].Contains("ADIF_GRUPO_") && 
                    //        !lista_carpetas[i].Contains(param.GetValue("excluir_lote", DateTime.Now, DateTime.Now)))
                    //    {
                    //        subruta = lista_carpetas[i].Replace(dirFTP, "");
                    //        subruta = subruta.Replace("/", "");
                    //        subruta = subruta.Replace("ADIF_", "") + "_CURVAS_ADIF";

                    //        rutas_validas.Add(dirFTP + "/" + lista_carpetas[i] + "/"
                    //            + subruta + "/"
                    //            + year + "/"
                    //            + subruta
                    //            + "_MES_" + mes_letra + "_" + year);
                    //    }

                    //}

                    pb.Show();
                    pb.progressBar.Step = 1;
                    pb.progressBar.Maximum = rutas_validas.Count();
                    pb.Text = "Comprobando archivos... ";

                    Cursor.Current = Cursors.WaitCursor;

                    lista_archivos = ftpClient.ListaDirectorios(dirFTP);

                    pb.progressBar.Value = jj;
                    percent = (jj / Convert.ToDouble(rutas_validas.Count())) * 100;
                    pb.txtDescripcion.Text = "Comprobando archivos en : " + rutas_validas[jj];
                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    pb.Refresh();

                    Application.DoEvents();

                    //for (int i = 0; i < rutas_validas.Count(); i++)
                    //{


                    //    lista_archivos = ftpClient.ListaDirectorios(rutas_validas[i]);
                    //    for (int j = 0; j < lista_archivos.Count(); j++)
                    //    {
                    //        if (lista_archivos[j].Contains("_CC1_"))
                    //            totalArchivos++;

                    //    }

                    //    pb.progressBar.Value = i;
                    //    percent = (i / Convert.ToDouble(rutas_validas.Count())) * 100;
                    //    pb.txtDescripcion.Text = "Comprobando archivos en : " + rutas_validas[i];
                    //    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    //    pb.Refresh();

                    //    Application.DoEvents();
                    //}

                    pb.Close();

                    if (totalArchivos > 0)
                    {

                        DirectoryInfo dir = new DirectoryInfo(param.GetValue("ftp_local_dir", DateTime.Now, DateTime.Now));
                        if (!dir.Exists)
                            dir.Create();

                        pb = new forms.FrmProgressBar();
                        pb.Show();
                        pb.progressBar.Step = 1;
                        pb.progressBar.Maximum = totalArchivos;
                        pb.Text = "Descargando " + totalArchivos + " archivos";
                        Cursor.Current = Cursors.WaitCursor;                        

                        for (int i = 0; i < rutas_validas.Count(); i++)
                        {
                            lista_archivos = ftpClient.ListaDirectorios(rutas_validas[i]);
                            for (int j = 0; j < lista_archivos.Count(); j++)
                            {
                                if (lista_archivos[j].Contains("_CC1_"))
                                {
                                    ftpClient.Download(rutas_validas[i] + "/" + lista_archivos[j],
                                        param.GetValue("ftp_local_dir", DateTime.Now, DateTime.Now) + lista_archivos[j]);

                                    totalArchivosDescargados++;

                                    pb.progressBar.Value = totalArchivosDescargados;
                                    percent = (totalArchivosDescargados / Convert.ToDouble(totalArchivos)) * 100;
                                    pb.txtDescripcion.Text = "Descargando: " + rutas_validas[i] + "/" + lista_archivos[j];                                       
                                    
                                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                                    pb.Refresh();

                                    Application.DoEvents();
                                }                                    
                                    
                            }
                            
                        }

                    }
                    Cursor.Current = Cursors.Default;
                    pb.Close();

                    MessageBox.Show("Descarga finalizada",
                   "Descarga FTP",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Information);

                }
                                             
            }
            catch (Exception er)
            {
                MessageBox.Show(er.Message,
                    "Descarga FTP",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void FrmAdif_Importar_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Medida", "FrmAdif_Importar", "N/A");
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
