using EndesaBusiness.cnmc;
using OfficeOpenXml.LoadFunctions.Params;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace GestionOperaciones.forms.contratacion
{
    public partial class FrmXML_A1550 : Form
    {
        EndesaBusiness.utilidades.Param p;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        EndesaBusiness.contratacion.eexxi.XML_GAS xml_gas;
        List<EndesaEntity.contratacion.gas.Informe_GAS_XML> lista;
        public FrmXML_A1550()
        {

            usage.Start("Contratación", "FrmXML_A1550", "N/A");
            InitializeComponent();

            p = new EndesaBusiness.utilidades.Param("gas_eexxi_xml_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmXML_A1550_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Contratación", "FrmXML_A1550", "N/A");
        }

        private void importarXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CargaXML();
        }

        private void CargaXML()
        {
            double percent = 0;

            int total_archivos = 0;
            int ficheros_procesados = 0;
            DialogResult result = DialogResult.Yes;
            int ii = 0;
            int num;

            EndesaBusiness.cnmc.CNMC cnmc;
            EndesaBusiness.sigame.SIGAME sigame;
            

            OpenFileDialog d = new OpenFileDialog();
            d.Title = p.GetValue("mensaje_ventana_xml", DateTime.Now, DateTime.Now);
            d.Filter = "zip files|*.zip";
            d.Multiselect = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                cnmc = new CNMC(DateTime.Now, DateTime.Now);
                sigame = new EndesaBusiness.sigame.SIGAME();

                forms.FrmProgressBar pb = new FrmProgressBar();

                if (!Directory.Exists(p.GetValue("inbox")))
                    Directory.CreateDirectory(p.GetValue("inbox"));

                BorrarContenidoDirectorio(p.GetValue("inbox", DateTime.Now, DateTime.Now));


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
                    //zip.Descomprimir(fileName, p.GetValue("inbox"));

                    EndesaBusiness.utilidades.ZIP zip = new EndesaBusiness.utilidades.ZIP();
                    zip.DescomprimirArchivoZip(fileName, p.GetValue("inbox"));
                }

                pb.Close();


                string[] files = Directory.GetFiles(p.GetValue("inbox", DateTime.Now, DateTime.Now), "*.xml");
                total_archivos = files.Count();

                p.UpdateParameter("fecha_ultima_carga", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                p = new EndesaBusiness.utilidades.Param("gas_eexxi_xml_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);

                #region Lee XML
                pb = new FrmProgressBar();
                pb.Text = "Carga archivos XML";
                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = files.Count();


                EndesaEntity.cnmc.gas.Proceso_A15_50 xml;
                EndesaBusiness.contratacion.eexxi.XML_GAS cont_xml =
                    new EndesaBusiness.contratacion.eexxi.XML_GAS();


                for (int i = 0; i < files.Count(); i++)
                {
                    FileInfo file_detail = new FileInfo(files[i]);
                    System.IO.StreamReader file = new System.IO.StreamReader(files[i]);
                    #region ProgressBar
                    percent = (i / Convert.ToDouble(files.Count())) * 100;
                    pb.progressBar.Increment(1);
                    pb.txtDescripcion.Text = "Importando " + i.ToString("#.###") + " / " + files.Count().ToString("#.###")
                        + " archivos --> " + file_detail.Name;
                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    pb.Refresh();
                    #endregion

                    XmlSerializer serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.gas.Proceso_A15_50));

                    System.Xml.Serialization.XmlSerializer reader =
                        new System.Xml.Serialization.XmlSerializer(typeof(EndesaEntity.cnmc.gas.Proceso_A15_50));

                    bool xml_valido = cont_xml.XML_Valido(file_detail.FullName);
                    if (xml_valido)
                    {
                        xml = (EndesaEntity.cnmc.gas.Proceso_A15_50)reader.Deserialize(file);
                        file.Close();

                        if (xml.a1550.tolltype != null)
                        {

                            /* Modificacion Victor 20231016
                                * Quitar filtro de tarifa
                            // Unicamente tratamos los XML que tienen el campo
                            // tolltype != R1, R2, R3, B*, S*

                            if (!xml.a1550.tolltype.Contains("B") &&
                                !xml.a1550.tolltype.Contains("S") &&
                                xml.a1550.tolltype != "R1" &&
                                xml.a1550.tolltype != "R2" &&
                                xml.a1550.tolltype != "R3"&&
                                int.TryParse(xml.a1550.tolltype, out num))
                            {
                                ficheros_procesados++;
                                cont_xml.Guarda_XML(file_detail.Name, xml);

                            }
                            * Fin Modificacion 20231016
                            */

                            /* ADDED Modificacion peticion JMG 20231120 */
                            if (cnmc.Es_Tarifa_GAS_XXI(xml.a1550.tolltype))
                            {
                                ficheros_procesados++;
                                cont_xml.Guarda_XML(file_detail.Name, xml);
                            }
                            else if (sigame.EsNIF_Vigente(xml.a1550.documentnum))
                            {
                                ficheros_procesados++;
                                cont_xml.Guarda_XML(file_detail.Name, xml);
                            }
                            /* END ADDED Modificacion peticion JMG 20231120 */

                        }
                    }


                }
                pb.Close();
                #endregion

                LoadData();

                MessageBox.Show("La importación ha concluido correctamente."
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "Se han procesado " + ficheros_procesados.ToString("N0") + " archivos de un total de " + files.Count().ToString("N0"),
                        "Importación ficheros XML",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);


                

            }
        }
        
       

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización XML A15 50";
            p.tabla = "gas_eexxi_xml_param";
            p.esquemaString = "CON";
            p.Show();
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

        private void FrmXML_A1550_Load(object sender, EventArgs e)
        {
            List<string> lista_drivers = new List<string>();

            EndesaBusiness.utilidades.Global utilGlobal = new EndesaBusiness.utilidades.Global();
            lista_drivers = utilGlobal.GetSystemDriverList();
            string driver = lista_drivers.Find(z => z == "Amazon Redshift (x64)");
            if (driver == null)
            {
                MessageBox.Show("No tiene instalado el driver Amazon RedShift (x64)"
                    + System.Environment.NewLine + "La aplicación no puede continuar y se cerrará.",
                   "ODBC RedShift NO ENCONTRADO",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Information);
                this.Close();
            }

            
            LoadData();
            
        }


        private void LoadData()
        {
            Cursor.Current = Cursors.WaitCursor;
            xml_gas = new EndesaBusiness.contratacion.eexxi.XML_GAS();

            //lista = xml_gas.Datos().OrderByDescending(z => z.fecha_inicio_xml).ToList(); 
            lista = xml_gas.Datos().Where(z => z.sistema == null).OrderByDescending(z => z.fecha_inicio_xml).ToList();

            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;
            lbl_total.Text = string.Format("Total Registros: {0:#,##0}", lista.Count());
            Cursor.Current = Cursors.Default;
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            xml_gas.InformeExcel(lista);
        }
    }
}
