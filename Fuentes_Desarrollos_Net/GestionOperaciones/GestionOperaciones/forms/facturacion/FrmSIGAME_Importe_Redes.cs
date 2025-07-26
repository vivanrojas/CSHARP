using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using System.Threading;
using System.Net.Http;
using OfficeOpenXml;
using System.IO;
using System.Globalization;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmSIGAME_Importe_Redes : Form
    {

        EndesaBusiness.facturacion.SIGAME_Importe_Redes importeRedes =
            new EndesaBusiness.facturacion.SIGAME_Importe_Redes();

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmSIGAME_Importe_Redes()
        {

            DateTime fd;
            DateTime fh;
            DateTime mesAnterior;
            int anio;
            int mes;
            int dias_del_mes;

            mesAnterior = DateTime.Now.AddMonths(-1);


            anio = mesAnterior.Year;
            mes = mesAnterior.Month;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            fh = new DateTime(anio, mes, dias_del_mes);
            fd = fh.AddMonths(-2);
            fd = new DateTime(fd.Year, fd.Month, 1);

            usage.Start("Facturación", "FrmSIGAME_Importe_Redes" ,"N/A");

            InitializeComponent();

            txtMaxFecha.Text = importeRedes.UltimaImportacion();

            dtFFACTDES.Value = fd;
            dtFFACTHAS.Value = fh;
                       
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
                
        private void btnBuscar_Click(object sender, EventArgs e)
        {
            bool usarFechaFactura;
            string cupsree = null;
               
            
            if (txtcupsree.Text != "")
                cupsree = txtcupsree.Text;

            usarFechaFactura = cmbTipoFecha.SelectedIndex == 0;
            

            importeRedes.CargaImporteConceptosEspeciales(usarFechaFactura, dtFFACTDES.Value, dtFFACTHAS.Value);
            importeRedes.CargaImportesConceptos(usarFechaFactura, dtFFACTDES.Value, dtFFACTHAS.Value, cupsree);
            importeRedes.CargaImporte_DTO(usarFechaFactura, dtFFACTDES.Value, dtFFACTHAS.Value);
            importeRedes.CargardgvCUPS(usarFechaFactura, dtFFACTDES.Value, dtFFACTHAS.Value, cupsree);

            dgvCups.DataSource = importeRedes.lista_facturas;
        }

        private void copyAlltoClipboard()
        {
            dgvCups.SelectAll();
            DataObject dataObj = dgvCups.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void btnExcel_Click(object sender, EventArgs e)
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
                importeRedes.ExportExcelOpenOffice(save.FileName);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }
        }

        private void CopyClipboard()
        {
            DataObject d = dgvCups.GetClipboardContent();
            Clipboard.SetDataObject(d);
        }

               
        private void salirToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Close();
        }

       

        private void btnExchange_Click(object sender, EventArgs e)
        {
            string dominio;
            string user;
            string password = "";

            Cursor = Cursors.WaitCursor;
            EndesaBusiness.Currency.ExchangeRate exchangeRate = new EndesaBusiness.Currency.ExchangeRate();

            dominio = System.Environment.UserDomainName;
            user = System.Environment.UserName;

            if (exchangeRate.necesitaCredenciales)
            {
                FrmCredencialesWeb f = new FrmCredencialesWeb();
                f.txtDominio.Text = System.Environment.UserDomainName;
                f.txtUser.Text = System.Environment.UserName;
                f.ShowDialog();
            }

            exchangeRate.DescargaEuroDolar(user, password, dominio);
            Cursor = Cursors.Default;
            txtMaxFecha.Text = importeRedes.UltimaImportacion();
           

        }

        

        private void FrmSIGAME_Importe_Redes_Load(object sender, EventArgs e)
        {
            this.cmbTipoFecha.SelectedIndex = 0;
        }

        private void parametrizaciónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización Importe Redes";
            p.tabla = "fact_p_exchange_param";
            p.esquemaString = "FAC";
            p.Show();
        }

        private void FrmSIGAME_Importe_Redes_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmSIGAME_Importe_Redes" ,"N/A");
        }
    }
}
