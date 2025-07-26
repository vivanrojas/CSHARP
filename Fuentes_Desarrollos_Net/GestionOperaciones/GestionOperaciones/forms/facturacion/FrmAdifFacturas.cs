using EndesaBusiness.facturacion.redshift;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Schema;
using static Microsoft.WindowsAPICodePack.Shell.PropertySystem.SystemProperties.System;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmAdifFacturas : Form
    {

        DateTime fd;
        DateTime fh;
        DateTime mesAnterior;
        int anio;
        int mes;
        int dias_del_mes;

        EndesaBusiness.facturacion.AdifFacturas adif_fact = new EndesaBusiness.facturacion.AdifFacturas();

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        

        int newSortColumn;
        ListSortDirection newColumnDirection = ListSortDirection.Ascending;
        public FrmAdifFacturas()
        {

            usage.Start("Facturación", "FrmAdifFacturas" ,"N/A");
            mesAnterior = DateTime.Now.AddMonths(-1);
            anio = mesAnterior.Year;
            mes = mesAnterior.Month;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            fh = new DateTime(anio, mes, dias_del_mes);
            fd = new DateTime(fh.Year, fh.Month, 1);

            InitializeComponent();

            dtFFACTDES.Value = fd;
            dtFFACTHAS.Value = fh;
            
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void CargardgvCUPS(string ffactdes, string ffacthas, string cupsree, string lote)
        {

            try
            {              
                
                Cursor.Current = Cursors.WaitCursor;
                //dgvCups.DataSource = adif_fact.CargaCups(ffactdes, ffacthas, cupsree, lote, chkDiff.Checked);
                Cursor.Current = Cursors.Default;                
                
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la importación de CUPS",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

        }

        //private void btnBuscar_Click(object sender, EventArgs e)
        //{
        //    this.CargardgvCUPS(dtFFACTDES.Value.ToString("yyyy-MM-dd"), dtFFACTHAS.Value.ToString("yyyy-MM-dd"), txtcupsree.Text, txtlote.Text);
        //}

        private void copyAlltoClipboard()
        {
            dgvCups.SelectAll();
            DataObject dataObj = dgvCups.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            adif_fact.ExportExcel();
            Cursor = Cursors.Default;
        }

        private void CopyClipboard()
        {
            DataObject d = dgvCups.GetClipboardContent();
            Clipboard.SetDataObject(d);
        }

        


        private void menuCopyPaste_Opening(object sender, CancelEventArgs e)
        {

        }

        private void importarToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
               

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void archivosFacturasADIFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "txt files|*.*";
            d.Multiselect = true;
            EndesaBusiness.facturacion.Adif_Ficheros fichero;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    fichero = new EndesaBusiness.facturacion.Adif_Ficheros(fileName);
                }

                MessageBox.Show("La importación ha concluido correctamente.",
                    "Importación ficheros facturas ADIF",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

               //this.CargardgvCUPS(dtFFACTDES.Value.ToString("yyyy-MM-dd"), dtFFACTHAS.Value.ToString("yyyy-MM-dd"), txtcupsree.Text, txtlote.Text);
                //EndesaBusiness.facturacion.Adif_fichero_factura ff = new adif.Adif_fichero_factura();
            }
        }

        private void archivosFacturasAdif_REG_toolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "txt files|*.txt";
            d.Multiselect = true;
            EndesaBusiness.facturacion.Adif_Ficheros fichero;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    fichero = new EndesaBusiness.facturacion.Adif_Ficheros(fileName);
                }

                MessageBox.Show("La importación ha concluido correctamente.",
                    "Importación ficheros regularización facturas ADIF",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                //this.CargardgvCUPS(dtFFACTDES.Value.ToString("yyyy-MM-dd"), dtFFACTHAS.Value.ToString("yyyy-MM-dd"), txtcupsree.Text, txtlote.Text);
                // EndesaEntity.facturacion.Adif_fichero_factura ff = new adif.Adif_fichero_factura();
            }
        }

        private void salirToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Close();
        }

        private void FrmAdifFacturas_Load(object sender, EventArgs e)
        {

        }

        private void FrmAdifFacturas_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmAdifFacturas" ,"N/A");
        }

        private void btnBuscar_Click_1(object sender, EventArgs e)
        {
            LoadData();
        }

        private void LoadData()
        {
            adif_fact = new EndesaBusiness.facturacion.AdifFacturas(Convert.ToDateTime(dtFFACTDES.Text), 
                Convert.ToDateTime(dtFFACTHAS.Text), txtcupsree.Text, txtlote.Text, chkDiff.Checked);

            Cursor.Current = Cursors.WaitCursor;
            dgvCups.AutoGenerateColumns = false;
            dgvCups.DataSource = adif_fact.lista_facturas;
            Cursor.Current = Cursors.Default;            

        }

        private List<EndesaEntity.facturacion.Adif_Factura> 
            OrdenaColumna(string columna, ListSortDirection direccion)
        {
            List<EndesaEntity.facturacion.Adif_Factura> l =
                new List<EndesaEntity.facturacion.Adif_Factura>();

            switch(columna)
            {
                case "CUPSREE":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.CUPSREE).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.CUPSREE).ToList();
                    break;

                case "LOTE":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.LOTE).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.LOTE).ToList();
                    break;

                case "medida_en_baja":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.medida_en_baja).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.medida_en_baja).ToList();
                    break;

                case "devolucion_de_energia":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.devolucion_de_energia).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.devolucion_de_energia).ToList();
                    break;

                case "cierres_energia":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.cierres_energia).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.cierres_energia).ToList();
                    break;

                case "existe_factura_sce":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.existe_factura_sce).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.existe_factura_sce).ToList();
                    break;

                case "existe_factura_adif":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.existe_factura_adif).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.existe_factura_adif).ToList();
                    break;

                case "CREFEREN":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.CREFEREN).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.CREFEREN).ToList();
                    break;

                case "SECFACTU":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.SECFACTU).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.SECFACTU).ToList();
                    break;

                case "FFACTURA":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.FFACTURA).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.FFACTURA).ToList();
                    break;

                case "FFACTDES":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.FFACTDES).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.FFACTDES).ToList();
                    break;

                case "FFACTHAS":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.FFACTHAS).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.FFACTHAS).ToList();
                    break;

                case "de_tfactura":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.de_tfactura).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.de_tfactura).ToList();
                    break;

                case "testfact":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.testfact).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.testfact).ToList();
                    break;

                case "CONSUMO_ADIF":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.CONSUMO_ADIF).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.CONSUMO_ADIF).ToList();
                    break;

                case "CONSUMO_SCE":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.CONSUMO_SCE).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.CONSUMO_SCE).ToList();
                    break;

                case "DIF_CONSUMO":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.DIF_CONSUMO).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.DIF_CONSUMO).ToList();
                    break;

                case "cnpr_adif":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.cnpr_adif).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.cnpr_adif).ToList();
                    break;

                case "cnpr_endesa":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.cnpr_endesa).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.cnpr_endesa).ToList();
                    break;

                case "DIF_TOTAL":
                    if (direccion == ListSortDirection.Ascending)
                        l = adif_fact.lista_facturas.OrderBy(z => z.DIF_TOTAL).ToList();
                    else
                        l = adif_fact.lista_facturas.OrderByDescending(z => z.DIF_TOTAL).ToList();
                    break;

            }

            return l;
        }

        private void dgvCups_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dgvCups.Columns[e.ColumnIndex].SortMode != DataGridViewColumnSortMode.NotSortable)
            {
                if (e.ColumnIndex == newSortColumn)
                {
                    if (newColumnDirection == ListSortDirection.Ascending)
                        newColumnDirection = ListSortDirection.Descending;
                    else
                        newColumnDirection = ListSortDirection.Ascending;
                }

                newSortColumn = e.ColumnIndex;

                dgvCups.AutoGenerateColumns = false;
                dgvCups.DataSource = OrdenaColumna(dgvCups.Columns[e.ColumnIndex].DataPropertyName, newColumnDirection); ;


            }
        }

        private void cierresDeEnergíaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmAdifCierresEnergia f = new medida.FrmAdifCierresEnergia();
            f.Show();
        }

        private void parámetrosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "adif_param";
            p.esquemaString = "FAC";
            p.cabecera = "Parámetros ADIF Facturación";
            p.Show();
        }
    }
}
