using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmPuntosSofisticados : Form
    {

        
        DataTable dt;
                
        EndesaBusiness.utilidades.Global g = new EndesaBusiness.utilidades.Global();

        EndesaBusiness.facturacion.InventarioSofisticadosManuales sofis = new EndesaBusiness.facturacion.InventarioSofisticadosManuales();

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmPuntosSofisticados()
        {
            usage.Start("Facturación", "FrmPuntosSofisticados" ,"N/A");
            InitializeComponent();

        }


        private void Cargadgv(string ccounips)
        {

            DateTime begin = new DateTime();
            try
            {                
                begin = DateTime.Now;
                

                
                if (txtCCOUNIPS.Text != "")                
                    sofis.Carga(txtCCOUNIPS.Text);                
                else                
                    sofis.Carga(null);               
                

                Cursor.Current = Cursors.WaitCursor;              



                dgvFacturas.AutoGenerateColumns = false;
                dgvFacturas.DataSource = sofis.lista;

                if (sofis.lista.Count == 0)
                {
                    MessageBox.Show("La consulta que ha realizado no devuelve ningún resultado.",
                    "La consulta no devuelte datos.",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                }

                Cursor.Current = Cursors.Default;
                               

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la búsqueda de Puntos Sofisticados",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }






        private void FrmFacturasOperaciones_Load(object sender, EventArgs e)
        {
            Cargadgv(txtCCOUNIPS.Text);
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            this.Cargadgv(txtCCOUNIPS.Text);
            Cursor = Cursors.Default;
        }

        private void toolTipNIF_Popup(object sender, PopupEventArgs e)
        {

        }





        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            DateTime ahora = new DateTime();
            int c = 0;

            Microsoft.Office.Interop.Excel.Application xlexcel;

            try
            {
                xlexcel = new Excel.Application();


                ahora = DateTime.Now;
                Cursor = Cursors.WaitCursor;

                #region Informe Normal

                copyAlltoClipboard();

                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                //
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlexcel.ActiveSheet.Cells(1, 2).Value = "CCOUNIPS";
                xlexcel.ActiveSheet.Cells(1, 3).Value = "DAPERSOC";
                xlexcel.ActiveSheet.Cells(1, 4).Value = "GRUPO";
                xlexcel.ActiveSheet.Cells(1, 5).Value = "Fecha Desde";
                xlexcel.ActiveSheet.Cells(1, 6).Value = "Fecha Hasta";
                xlexcel.ActiveSheet.Cells(1, 7).Value = "PRECIOS";
                xlexcel.ActiveSheet.Cells(1, 8).Value = "FACTURAS A CUENTA";

                c = 5;
                for (int i = 1; i <= c; i++)
                {
                    xlexcel.ActiveSheet.Cells(1, i).Font.Bold = true;
                }

                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[2, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                #endregion




            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message,
                "Error en la construcción de la consulta",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
            Cursor = Cursors.Default;

        }

        private void copyAlltoClipboard()
        {
            dgvFacturas.SelectAll();
            DataObject dataObj = dgvFacturas.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void listBoxEmpresas_SelectedIndexChanged(object sender, EventArgs e)
        {

        }







        private void listBoxLinea_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void txtCod_TextChanged(object sender, EventArgs e)
        {

        }



        private void btnAdd_Click(object sender, EventArgs e)
        {

            EndesaBusiness.facturacion.InventarioSofisticadosManuales sofis
                = new EndesaBusiness.facturacion.InventarioSofisticadosManuales();
            forms.facturacion.FormPuntosSofisticados_Edit f
                = new FormPuntosSofisticados_Edit();
            f.Text = "Añadir nuevo registro";
            f.ShowDialog(this);

            if (!f.pressCancel)
            {
                if (f.txtCCOUNIPS.Text != null &&
                f.txtFD.Text != null)
                {
                    //sofis?.cups13 = "";
                    sofis.cups13 = f.txtCCOUNIPS.Text;
                    sofis.cups20 = f.txtcusp20.Text;
                    sofis.dapersoc = f.txtDAPERSOC.Text;
                    sofis.grupo = f.txtGRUPO.Text;
                    sofis.fd = Convert.ToDateTime(f.txtFD.Text);
                    sofis.fh = Convert.ToDateTime(f.txtFH.Text);
                    sofis.precios = f.txtPrecios.Text;
                    sofis.facturas_a_cuenta = f.chkFACTURAS_A_CUENTA.Checked;
                    sofis.Add();
                    Cargadgv(txtCCOUNIPS.Text);

                    // string a = $"hola {sofis.cups13 } kkk ";

                }

            }




        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            EndesaBusiness.facturacion.InventarioSofisticadosManuales sofis =
                new EndesaBusiness.facturacion.InventarioSofisticadosManuales();

            try
            {
                DataGridViewRow row = new DataGridViewRow();
                int fila = dgvFacturas.CurrentRow.Index;

                forms.facturacion.FormPuntosSofisticados_Edit f
                    = new FormPuntosSofisticados_Edit();
                f.Text = "Editar registro";
                f.txtCCOUNIPS.Enabled = false;
                f.txtFD.Enabled = false;

                row = dgvFacturas.Rows[fila];
                sofis.cups13 = row.Cells[0].Value.ToString();
                sofis.cups20 = row.Cells[1].Value.ToString();

                if (row.Cells[2].Value.ToString() != "")
                    sofis.dapersoc = row.Cells[2].Value.ToString();
                sofis.grupo = row.Cells[3].Value.ToString();
                sofis.fd = Convert.ToDateTime(row.Cells[4].Value);
                if (row.Cells[5].Value.ToString() != "")
                    sofis.fh = Convert.ToDateTime(row.Cells[5].Value);
                sofis.precios = row.Cells[6].Value.ToString();
                sofis.facturas_a_cuenta = (row.Cells[7].Value.ToString() == "S" ? true : false);

                f.txtCCOUNIPS.Text = sofis.cups13;
                f.txtcusp20.Text = sofis.cups20;
                f.txtDAPERSOC.Text = sofis.dapersoc;
                f.txtGRUPO.Text = sofis.grupo;
                f.txtFD.Text = sofis.fd.ToString();
                f.txtFH.Text = sofis.fh.ToString();
                f.txtPrecios.Text = sofis.precios;
                f.chkFACTURAS_A_CUENTA.Checked = sofis.facturas_a_cuenta;

                f.ShowDialog(this);

                if (f.txtGRUPO.Text != sofis.grupo ||
                    f.txtDAPERSOC.Text != sofis.dapersoc ||
                    Convert.ToDateTime(f.txtFH.Text) != sofis.fh ||
                     f.chkFACTURAS_A_CUENTA.Checked != sofis.facturas_a_cuenta)
                {

                    sofis.dapersoc = f.txtDAPERSOC.Text;
                    sofis.grupo = f.txtGRUPO.Text;
                    sofis.fh = Convert.ToDateTime(f.txtFH.Text);
                    sofis.precios = f.txtPrecios.Text;
                    sofis.facturas_a_cuenta = f.chkFACTURAS_A_CUENTA.Checked;

                    sofis.Update();
                    Cargadgv(txtCCOUNIPS.Text);

                }

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message,
               "Error en la edición de los datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }

        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            
            DialogResult r;

            try
            {
                DataGridViewRow row = new DataGridViewRow();
                int fila = dgvFacturas.CurrentRow.Index;
                row = dgvFacturas.Rows[fila];
                sofis.cups13 = row.Cells[0].Value.ToString();
                sofis.fd = Convert.ToDateTime(row.Cells[4].Value);

                r = MessageBox.Show("¿Desea borrar el registros con CCOUNIPS " + sofis.cups13 + "?",
                    "Borrar registro con código " + sofis.cups13,
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (r == DialogResult.Yes)
                {
                    sofis.Del();
                    Cargadgv(txtCCOUNIPS.Text);
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message,
               "Error en el borrado de los datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        private void FrmPuntosSofisticados_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmPuntosSofisticados" ,"N/A");
        }
    }
}
