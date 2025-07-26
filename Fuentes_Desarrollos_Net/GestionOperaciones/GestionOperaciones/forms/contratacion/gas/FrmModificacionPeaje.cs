using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.contratacion.gas
{
    public partial class FrmModificacionPeaje : Form
    {

        List<EndesaEntity.contratacion.gas.Cups_Peaje> lc = new List<EndesaEntity.contratacion.gas.Cups_Peaje>();
        EndesaBusiness.contratacion.gestionATRGas.GestionATRGas gestionATR;
        EndesaBusiness.utilidades.Param pp;
        EndesaBusiness.cnmc.XML formato_xml = new EndesaBusiness.cnmc.XML();
        EndesaBusiness.cnmc.CNMC cnmc = new EndesaBusiness.cnmc.CNMC(new DateTime(2021,10,01), new DateTime(2021, 10, 01));
        List<EndesaEntity.contratacion.gas.Cups_Peaje> lista;


        EndesaBusiness.contratacion.gestionATRGas.Distribuidoras distribuidoras;
        EndesaBusiness.sigame.SIGAME inventario_gas;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        public bool hayError { get; set; }
        public string descripcion_error { get; set; }

        public FrmModificacionPeaje()
        {
            usage.Start("Contratación", "FrmModificacionPeaje" ,"N/A");
            InitializeComponent();
        }

        private void FrmModificacionPeaje_Load(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            // Inventario distribuidoras
            distribuidoras = new EndesaBusiness.contratacion.gestionATRGas.Distribuidoras(true);

            // Inventario de puntos GAS
            inventario_gas = new EndesaBusiness.sigame.SIGAME();

            pp = new EndesaBusiness.utilidades.Param("atrgas_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);

            gestionATR = new EndesaBusiness.contratacion.gestionATRGas.GestionATRGas();
            Cursor.Current = Cursors.Default;

        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnGenerarXML_Click(object sender, EventArgs e)
        {

            string secuencial;
            string fileName = "";
            string fechaHora = "";
            string destinycompany = "";
            string distri = "";
            Dictionary<string, List<EndesaEntity.contratacion.gas.Cups_Peaje>> dic_distri =
                new Dictionary<string, List<EndesaEntity.contratacion.gas.Cups_Peaje>>();

            EndesaBusiness.utilidades.ZIP zip = new EndesaBusiness.utilidades.ZIP();

            EndesaBusiness.cnmc.XML formato_xml = new EndesaBusiness.cnmc.XML();

            DialogResult result = MessageBox.Show("¿Desea generar los XML de las solicitudes?"                  
                    , "Generar solicitudes XML",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                //secuencial = Convert.ToInt32(pp.GetValue("secuencial_solicitud", DateTime.Now, DateTime.Now));
                secuencial = DateTime.Now.ToString("yyMMddHHmmss").ToString();
                Thread.Sleep(1000);

                if (!Directory.Exists(pp.GetValue("inbox", DateTime.Now, DateTime.Now)))
                    Directory.CreateDirectory(pp.GetValue("inbox", DateTime.Now, DateTime.Now));


                foreach (EndesaEntity.contratacion.gas.Cups_Peaje p in lista)
                {
                     
                    distri = distribuidoras.Codigo_XML_CNMC_Distribuidora(inventario_gas.Distribuidora(p.cups).ToUpper());
                    List<EndesaEntity.contratacion.gas.Cups_Peaje> o;
                    if (!dic_distri.TryGetValue(distri, out o))
                    {
                        o = new List<EndesaEntity.contratacion.gas.Cups_Peaje>();
                        o.Add(p);
                        dic_distri.Add(distri, o);
                    }
                    else                    
                        o.Add(p);
                    

                }

                foreach(KeyValuePair<string, List<EndesaEntity.contratacion.gas.Cups_Peaje>> c in dic_distri)
                {

                    destinycompany = c.Key;

                    foreach (EndesaEntity.contratacion.gas.Cups_Peaje p in c.Value)
                    {

                        //secuencial++;
                        EndesaEntity.cnmc.XML_A1_43 xml_a1_43 = new EndesaEntity.cnmc.XML_A1_43();
                        //xml_a1_43.comreferencenum = pp.GetValue("prefijo_solicitud", DateTime.Now, DateTime.Now)
                        //       + DateTime.Now.ToString("yyyy") + secuencial.ToString().PadLeft(4, '0');

                        xml_a1_43.comreferencenum = secuencial;

                        xml_a1_43.destinycompany =
                                distribuidoras.Codigo_XML_CNMC_Distribuidora(inventario_gas.Distribuidora(p.cups).ToUpper());

                        destinycompany = xml_a1_43.destinycompany;
                        xml_a1_43.dispatchingcompany = "0007";
                        xml_a1_43.documentnum = inventario_gas.NIF(p.cups);
                        xml_a1_43.cups = p.cups;
                        xml_a1_43.newtolltype = cnmc.Codigo_Tipo_Peaje(p.peaje);


                        fechaHora = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
                        fileName = pp.GetValue("ubicacion_xml_modificacion_peajes", DateTime.Now, DateTime.Now)
                            + @"\" + pp.GetValue("prefijo_archivo_a1_05", DateTime.Now, DateTime.Now)                            
                            + fechaHora
                            + ".xml";


                        FileInfo file = new FileInfo(fileName);
                        formato_xml.CreaXML_A1_05(file, xml_a1_43);
                        gestionATR.GuardaNumSecuencialTemporal(secuencial);
                    }

                    // Comprimimos los XML 
                    fechaHora = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    fileName = pp.GetValue("ubicacion_xml_modificacion_peajes", DateTime.Now, DateTime.Now) + @"\"
                            + pp.GetValue("prefijo_archivo_a1_05", DateTime.Now, DateTime.Now)
                            + destinycompany + "_"
                            + fechaHora
                            + ".zip";
                    FileInfo archivo = new FileInfo(fileName);

                    zip.ComprimirVarios(pp.GetValue("ubicacion_xml_modificacion_peajes", DateTime.Now, DateTime.Now) + @"\",
                        ".*\\.(xml)$", archivo.FullName);

                    BorrarContenidoDirectorio(pp.GetValue("ubicacion_xml_modificacion_peajes", DateTime.Now, DateTime.Now));

                }




                

                MessageBox.Show("Modificación Generada correctamente.",                
                     "ZIP(s) Completado(s)",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Information);
            }

        }

        

        private void importaciónExcelVentasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = false;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    lista = CargaExcel(fileName);

                    if (!hayError)
                    {                        
                        Carga_DGV(lista);
                        btnGenerarXML.Enabled = true;
                    }
                    else
                    {
                        MessageBox.Show(descripcion_error,
                           "Importar Excel",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Information);
                    }
                }
            }
        }

        private List<EndesaEntity.contratacion.gas.Cups_Peaje> CargaExcel(string fichero)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            bool firstOnly = true;
            string cabecera = "";

           


            try
            {
                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();
                List<EndesaEntity.contratacion.gas.Cups_Peaje> lista = new List<EndesaEntity.contratacion.gas.Cups_Peaje>();


                f = 1; // Porque la primera fila es la cabecera
                for (int i = 0; i < 5000; i++)
                {
                    c = 1;
                    if (firstOnly)
                    {
                        cabecera = workSheet.Cells[1, 1].Value.ToString()
                           + workSheet.Cells[1, 2].Value.ToString()
                           + workSheet.Cells[1, 3].Value.ToString()
                           + workSheet.Cells[1, 4].Value.ToString();



                        if (!EstructuraCorrecta(cabecera))
                        {
                            this.hayError = true;
                            this.descripcion_error = "La estructura del archivo excel no es la correcta.";
                            break;
                        }

                        firstOnly = false;
                    }

                    f++;

                    if (workSheet.Cells[f, 1].Value == null)
                        break;

                    if (workSheet.Cells[f, 1].Value.ToString()
                        + workSheet.Cells[f, 2].Value.ToString()
                        + workSheet.Cells[f, 3].Value.ToString()
                        + workSheet.Cells[f, 4].Value.ToString() == "")
                    {
                        break;
                    }
                    else
                    {
                        if(workSheet.Cells[f, 6].Value == null)
                        {

                            EndesaEntity.contratacion.gas.Cups_Peaje cups =
                            new EndesaEntity.contratacion.gas.Cups_Peaje();

                            cups.cups = workSheet.Cells[f, 3].Value.ToString(); 
                            if(cups.cups.Trim().Length == 20)
                            {
                                cups.nif = inventario_gas.NIF(cups.cups.Trim());
                                cups.cliente = inventario_gas.NombreCliente(cups.cups.Trim());
                                cups.distribuidora = inventario_gas.Distribuidora(cups.cups.Trim());
                                cups.fecha_efecto = new DateTime(2021, 10, 01).Date;
                                cups.peaje = workSheet.Cells[f, 4].Value.ToString();
                            }

                            cups.peaje = workSheet.Cells[f, 4].Value.ToString();

                            lista.Add(cups);
                        }
                    }


                }

                return lista;
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
                           "Importar Excel",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error);

                return null;
            }
        }

        private void BorrarContenidoDirectorio(string directorio)
        {
            string[] listaArchivos;
            FileInfo file;

            listaArchivos = Directory.GetFiles(directorio,"A1_05_SCTD*.xml");
            for (int i = 0; i < listaArchivos.Count(); i++)
            {
                file = new FileInfo(listaArchivos[i]);
                file.Delete();
            }
        }

        private bool EstructuraCorrecta(string cabecera)
        {


            if (cabecera.ToUpper().Trim() == "NIFCLIENTECUPSNUEVO PEAJE")
                return true;
            else
                return false;
        }

        private void Carga_DGV(List<EndesaEntity.contratacion.gas.Cups_Peaje> lista)
        {
            lblTotalRegistros.Text = string.Format("Total Registros: {0:#,##0}", lista.Count);
            dgvd.AutoGenerateColumns = false;
            dgvd.DataSource = lista;
        }

        private void FrmModificacionPeaje_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Contratación", "FrmModificacionPeaje" ,"N/A");
        }
    }
}
