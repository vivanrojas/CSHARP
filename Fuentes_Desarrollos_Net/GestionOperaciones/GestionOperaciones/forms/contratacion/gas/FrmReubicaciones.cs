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
using System.Xml;
using EndesaBusiness.cnmc;
using EndesaEntity.facturacion.mes13;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace GestionOperaciones.forms.contratacion.gas
{
    public partial class FrmReubicaciones : Form
    {

        EndesaBusiness.utilidades.Param p;
        EndesaBusiness.contratacion.gestionATRGas.Reubicaciones reubicaciones;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmReubicaciones()
        {

            usage.Start("Contratación", "FrmReubicaciones" ,"N/A");
            InitializeComponent();

            txt_fecha_hasta.Value = new DateTime(DateTime.Now.Year, 10, 1);
            txt_fecha_desde.Value = txt_fecha_hasta.Value.AddYears(-1);

            p = new EndesaBusiness.utilidades.Param("atrgas_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
            
        }

        private void FrmReubicaciones_Load(object sender, EventArgs e)
        {

            LoadData();
            LoadDataCargas();

        }

        private void LoadData()
        {
            Cursor.Current = Cursors.WaitCursor;
            reubicaciones = new EndesaBusiness.contratacion.gestionATRGas.Reubicaciones(txt_fecha_desde.Value, txt_fecha_hasta.Value, chk_all.Checked);
            dgv.AutoGenerateColumns = false;
            dgv.DataSource = reubicaciones.informe;
            lbl_total_reubicaciones.Text = string.Format("Total Registros: {0:#,##0}", dgv.RowCount);
            

        }


        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void importarXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int ii = 0;
            double percent = 0;
            int total_archivos = 0;
            int archivos_SIGAME = 0;
            DateTime max_fecha_solicitud = new DateTime();
            DateTime min_fecha_solicitud = new DateTime();

            OpenFileDialog d = new OpenFileDialog();            

            d.Title = "Carga de ZIP XML Reubicaciones";
            d.Filter = "zip files|*.zip";
            d.Multiselect = true;

            max_fecha_solicitud = DateTime.MinValue;
            min_fecha_solicitud = new DateTime(4999, 12, 31);

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                reubicaciones = new EndesaBusiness.contratacion.gestionATRGas.Reubicaciones(txt_fecha_desde.Value, txt_fecha_hasta.Value, false);

                forms.FrmProgressBar pb = new FrmProgressBar();

                if (!Directory.Exists(p.GetValue("inbox_reubicaciones")))
                    Directory.CreateDirectory(p.GetValue("inbox_reubicaciones"));

                if (!Directory.Exists(p.GetValue("RutaSalidaXML_Reubicaciones_ERROR")))
                    Directory.CreateDirectory(p.GetValue("RutaSalidaXML_Reubicaciones_ERROR"));


                BorrarContenidoDirectorio(p.GetValue("inbox_reubicaciones"));

                pb.Text = "Descomprimiendo ...";
                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = Convert.ToInt32(d.FileNames.Count());

                foreach (string fileName in d.FileNames)
                {
                    ii++;
                    percent = (ii / Convert.ToDouble(d.FileNames.Count())) * 100;                    
                    pb.progressBar.Increment(1);
                    pb.txtDescripcion.Text = "Extrayendo " + ii.ToString("#,##0") + " / " +
                        d.FileNames.Count().ToString("#,##0") + " archivos --> " + fileName;
                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    pb.Refresh();

                    //EndesaBusiness.utilidades.ZipUnZip zip = new EndesaBusiness.utilidades.ZipUnZip();
                    //zip.Descomprimir(fileName, p.GetValue("inbox_reubicaciones", DateTime.Now, DateTime.Now));
                    EndesaBusiness.utilidades.ZIP zip = new EndesaBusiness.utilidades.ZIP();
                    zip.DescomprimirArchivoZip(fileName, p.GetValue("inbox_reubicaciones"));

                }

                pb.Close();

                string[] files = Directory.GetFiles(p.GetValue("inbox_reubicaciones", DateTime.Now, DateTime.Now), "*.xml");
                total_archivos = files.Count();

                pb = new FrmProgressBar();
                pb.Text = "Analizando archivos XML";
                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = files.Count();               


                for (int i = 0; i < files.Count(); i++)
                {
                    ii++;
                    percent = (ii / Convert.ToDouble(total_archivos)) * 100;
                    pb.progressBar.Increment(1);
                    pb.txtDescripcion.Text = "Analizando " + ii.ToString("#,##0") + " / " +
                        total_archivos.ToString("#,##0") + " archivos --> " + files[i];
                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    pb.Refresh();

                    FileInfo file = new FileInfo(files[i]);
                                       
                    System.Xml.Serialization.XmlSerializer reader =
                        new System.Xml.Serialization.XmlSerializer(typeof(EndesaEntity.cnmc.gas.V25_2019_12_17.Proceso_A12_26));
                    System.IO.StreamReader file_xml = new System.IO.StreamReader(file.FullName);
                    EndesaEntity.cnmc.gas.V25_2019_12_17.Proceso_A12_26 xml = 
                        (EndesaEntity.cnmc.gas.V25_2019_12_17.Proceso_A12_26)reader.Deserialize(file_xml);
                    file_xml.Close();

                    if(xml.a1226.reqdate > max_fecha_solicitud)
                        max_fecha_solicitud = xml.a1226.reqdate;


                    if (xml.a1226.reqdate < min_fecha_solicitud)
                        min_fecha_solicitud = xml.a1226.reqdate;

                    if (reubicaciones.ExisteCUPS(xml.a1226.cups))
                    {
                        archivos_SIGAME++;
                        reubicaciones.GuardaXML(file.Name, xml);
                    }

                }

                pb.Close();
                Cursor.Current = Cursors.Default;

                reubicaciones.InsertaCarga(total_archivos.ToString("#,##0"),
                    min_fecha_solicitud.ToString("dd/MM/yyyy"),
                    max_fecha_solicitud.ToString("dd/MM/yyyy"));

                MessageBox.Show("La importación ha concluido correctamente."
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "Se han procesado " + (archivos_SIGAME).ToString("#,##0")
                        + " archivos de " + total_archivos.ToString("#,##0"),
                        "Importación ficheros XML",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                LoadData();
                LoadDataCargas();

            }
        }

        private void importarXMLToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            int ii = 0;
            double percent = 0;
            int total_archivos = 0;
            int archivos_SIGAME = 0;

            CommonOpenFileDialog d = new CommonOpenFileDialog();
            d.IsFolderPicker = true;
            d.Title = "Carga de directorio XML Reubicaciones";

            DateTime min_fecha_solicitud = new DateTime();
            DateTime max_fecha_solicitud = new DateTime();

            min_fecha_solicitud = DateTime.MaxValue;
            max_fecha_solicitud = DateTime.MinValue;

            if (d.ShowDialog() == CommonFileDialogResult.Ok)
            {
                forms.FrmProgressBar pb = new FrmProgressBar();                

                foreach (string dirName in d.FileNames)
                {

                    string[] files = Directory.GetFiles(dirName, "*.xml");
                    total_archivos = files.Count();

                    pb = new FrmProgressBar();
                    pb.Text = "Analizando archivos XML";
                    pb.Show();
                    pb.progressBar.Step = 1;
                    pb.progressBar.Maximum = total_archivos;

                    for (int i = 0; i < files.Count(); i++)
                    {
                        ii++;
                        percent = (ii / Convert.ToDouble(total_archivos)) * 100;
                        pb.progressBar.Increment(1);
                        pb.txtDescripcion.Text = "Analizando " + ii.ToString("#,##0") + " / " +
                            total_archivos.ToString("#,##0") + " archivos --> " + files[i];
                        pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                        pb.Refresh();

                        FileInfo file = new FileInfo(files[i]);                        

                        System.Xml.Serialization.XmlSerializer reader =
                            new System.Xml.Serialization.XmlSerializer(typeof(EndesaEntity.cnmc.gas.V25_2019_12_17.Proceso_A12_26));
                        System.IO.StreamReader file_xml = new System.IO.StreamReader(file.FullName);
                        EndesaEntity.cnmc.gas.V25_2019_12_17.Proceso_A12_26 xml =
                            (EndesaEntity.cnmc.gas.V25_2019_12_17.Proceso_A12_26)reader.Deserialize(file_xml);
                        file_xml.Close();

                        if (reubicaciones.ExisteCUPS(xml.a1226.cups))
                        {
                            archivos_SIGAME++;
                            reubicaciones.GuardaXML(file.Name, xml);
                        }

                        if (xml.a1226.reqdate < min_fecha_solicitud)
                            min_fecha_solicitud = xml.a1226.reqdate;


                        if (xml.a1226.reqdate > max_fecha_solicitud)
                            max_fecha_solicitud = xml.a1226.reqdate;
                    }


                    Console.WriteLine();
                }

                pb.Close();

                reubicaciones.InsertaCarga(total_archivos.ToString("#,##0"),
                    min_fecha_solicitud.ToString("dd/MM/yyyy"),
                   max_fecha_solicitud.ToString("dd/MM/yyyy"));

                MessageBox.Show("La importación ha concluido correctamente."
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "Se han procesado " + (archivos_SIGAME).ToString("#,##0") 
                        + " archivos de " + total_archivos.ToString("#,##0"),
                        "Importación ficheros XML",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                LoadData();

            }
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void LoadDataCargas()
        {
            Cursor.Current = Cursors.Default;
            reubicaciones.CargaCargas();
            dgvCargas.AutoGenerateColumns = false;
            dgvCargas.DataSource = reubicaciones.cargas;
            lbl_total_registros_cargas.Text = string.Format("Total Registros: {0:#,##0}", dgvCargas.RowCount);
        }

        private void cmdExcel_Click(object sender, EventArgs e)
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
                reubicaciones.GeneraInformeExcel(save.FileName);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            reubicaciones.InsertaRegistroSinArchivos();
            Cursor.Current = Cursors.Default;
            LoadDataCargas();
        }

        private void dgvCargas_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            //if (e.RowIndex != -1 && (e.ColumnIndex == 0 || e.ColumnIndex == 1))
            if (e.RowIndex != -1)
            {
                
                int fila = this.dgvCargas.CurrentCell.RowIndex;
                int columna = this.dgvCargas.CurrentCell.ColumnIndex;

                foreach (DataGridViewRow row in dgvCargas.Rows)
                {

                    if (row.Index == fila)
                    {
                        DataGridViewCell fecha_carga = (DataGridViewCell)
                        row.Cells[0];
                        DataGridViewCell fecha_min = (DataGridViewCell)
                           row.Cells[2];
                        DataGridViewCell fecha_max = (DataGridViewCell)
                           row.Cells[3];


                        reubicaciones.UpdateFecha(Convert.ToDateTime(fecha_carga.Value),
                            Convert.ToString(fecha_min.Value), Convert.ToString(fecha_max.Value));
                    }
                        
                }   
            }

              

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

        private void modificarFechasToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void FrmReubicaciones_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Contratación", "FrmReubicaciones" ,"N/A");
        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "ps_param";
            p.esquemaString = "CON";
            p.cabecera = "Parámetros PS";
            p.Show();
        }
    }
}
