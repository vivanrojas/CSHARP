using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmFacturasOperaciones : Form
    {

        DateTime fd;
        DateTime fh;
        int anio;
        int mes;
        int dias_del_mes;
        DataTable dt;

        String mainQuery = null;

        String conceptos = null;
        String cupsree = null;
        String tipoNegocio = null;
        String tiposFactura = null;
        String cnifdnic = null;
        String empresas = null;
        String cfactura = null;

        //List<string> listacups = new List<string>();
        //List<string> listanifs = new List<string>();
        string listacups;
        string listanifs;
        List<string> lista_empresas;


        EndesaBusiness.facturacion.Inffact inffact = new EndesaBusiness.facturacion.Inffact();

        EndesaBusiness.utilidades.Global g = new EndesaBusiness.utilidades.Global();

        // facturacion.EmpresasTitulares let = new facturacion.EmpresasTitulares();
        EndesaBusiness.facturacion.SegmentosMercado let =
            new EndesaBusiness.facturacion.SegmentosMercado();
        EndesaBusiness.facturacion.TiposFactura ltf =
            new EndesaBusiness.facturacion.TiposFactura();

        
        List<EndesaEntity.facturacion.FacturaInffact> lf =
            new List<EndesaEntity.facturacion.FacturaInffact>();


        bool otrosConceptos;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmFacturasOperaciones()
        {

            usage.Start("Facturación", "FrmFacturasOperaciones" ,"N/A");

            InitializeComponent();
            this.InitListBox();
            this.CargadgvConceptos();
            this.cmbTipoFecha.SelectedIndex = 0;
            this.cmbTipoInforme.SelectedIndex = 0;

            anio = DateTime.Now.Year;
            mes = (DateTime.Now.Month);
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            fh = new DateTime(anio, mes, dias_del_mes);
            fd = new DateTime(fh.Year, fh.Month, 1);

            txtFFACTURADES.Value = fd;
            txtFFACTURAHAS.Value = fh;

            lista_empresas = new List<string>();
                     
        }

        private void InitListBox()
        {

            for (int i = listBoxLinea.Items.Count - 1; i == 0; i--)
            {
                listBoxLinea.Items.RemoveAt(i);
            }

            listBoxLinea.Items.Add("LUZ");
            listBoxLinea.Items.Add("GAS");

            for (int i = listBoxEmpresas.Items.Count - 1; i == 0; i--)
            {
                listBoxEmpresas.Items.RemoveAt(i);
            }

            for (int i = 1; i <= let.NumRegistros; i++)
            {
                let.GetPosicionID(i);
                listBoxEmpresas.Items.Add(let.et.descripcion);
            }

            for (int i = listBoxTiposFactura.Items.Count - 1; i == 0; i--)
            {
                listBoxTiposFactura.Items.RemoveAt(i);
            }


            foreach(KeyValuePair<int, EndesaEntity.facturacion.Tipos_Factura_Tabla> p in ltf.dic)            
                listBoxTiposFactura.Items.Add(p.Value.descripcion);
            

            

        }

        private void CargadgvConceptos()
        {
            Cursor.Current = Cursors.WaitCursor;                
            dgvConceptos.AutoGenerateColumns = false;
            dgvConceptos.DataSource = inffact.ConceptosFacturacion(lista_empresas, txtNum.Text, txtCod.Text, txtDescripcion.Text);
            Cursor.Current = Cursors.Default;            
        }

    private void CargadgvFacturas()
        {
                       
            Boolean firstOnly = true;            
            Boolean usarFechaFactura;
            DateTime begin = new DateTime();

            try
            {
                lf.Clear();
                if (txtFFACTURADES.Value > txtFFACTURAHAS.Value)
                {
                    MessageBox.Show("La fecha Desde no puede ser superior a la fecha Hasta." +
                        System.Environment.NewLine + "Por favor reintente."
                        , "Error en fechas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtFFACTURAHAS.Focus();

                }
                else
                {

                    begin = DateTime.Now;
                    usarFechaFactura = cmbTipoFecha.SelectedIndex == 0;

                    

                    #region cfactura
                    if (txtCFACTURA.Text != "")
                    {
                        cfactura = txtCFACTURA.Text;
                    }
                    else
                    {
                        cfactura = null;
                    }
                    #endregion

                    #region cupsree
                    if (txtCUPSREE.Text != "")                    
                        cupsree = txtCUPSREE.Text;
                    else                    
                        cupsree = null;


                    firstOnly = true;
                    for (int i = 0; i < txtListaCUPS.Lines.Length ; i++)
                    {
                        if (firstOnly)
                        {
                            cupsree = txtListaCUPS.Lines[i];
                            firstOnly = false;
                        }else
                        {
                            if (txtListaCUPS.Lines[i].Trim() != "")
                                cupsree = cupsree + "|" + txtListaCUPS.Lines[i];
                        }
                    }

                    #endregion

                    #region cnifdnic
                    ;
                    if (txtCNIFDNIC.Text != "")                   
                        cnifdnic = txtCNIFDNIC.Text;                   
                    else                  
                        cnifdnic = null;

                    firstOnly = true;
                    for (int i = 0; i < txtListaNIFS.Lines.Length; i++)
                    {
                        if (firstOnly)
                        {
                            cnifdnic = txtListaNIFS.Lines[i];
                            firstOnly = false;
                        }
                        else
                        {
                            if (txtListaNIFS.Lines[i].Trim() != "")
                                cnifdnic = cnifdnic + "|" + txtListaNIFS.Lines[i];
                        }
                    }

                    #endregion

                    #region tipoNegocio
                    if (listBoxLinea.SelectedItems.Count > 0)
                    {
                        foreach (Object values in listBoxLinea.SelectedItems)
                        {
                            if (firstOnly)
                            {
                                tipoNegocio = values.ToString().Substring(0, 1);
                                firstOnly = false;
                            }
                            else
                            {
                                tipoNegocio = tipoNegocio + ";" + values.ToString().Substring(0, 1);
                            }
                        }
                        firstOnly = true;
                    }
                    else
                    {
                        tipoNegocio = null;
                    }
                    #endregion

                    #region tiposFactura
                    if (listBoxTiposFactura.SelectedItems.Count > 0)
                    {
                        foreach (Object values in listBoxTiposFactura.SelectedItems)
                        {
                            if (firstOnly)
                            {
                                tiposFactura = values.ToString();
                                firstOnly = false;
                            }
                            else
                            {
                                tiposFactura = tiposFactura + ";" + values.ToString();
                            }
                        }

                        firstOnly = true;
                    }
                    else
                    {
                        tiposFactura = null;
                    }
                    #endregion

                    #region empresas

                    firstOnly = true;
                    if (listBoxEmpresas.SelectedItems.Count > 0)
                    {                        
                        foreach (Object values in listBoxEmpresas.SelectedItems)
                        {
                            if (firstOnly)
                            {
                                empresas = values.ToString();
                                firstOnly = false;
                            }
                            else
                            {
                                empresas = empresas + ";" + values.ToString();
                            }
                        }

                        firstOnly = true;

                    }
                    else
                    {
                        empresas = null;
                    }
                    #endregion

                    #region conceptos
                    
                    firstOnly = true;
                    conceptos = null;
                    for (int i = 0; i < txtListaConceptos.Lines.Length; i++)
                    {
                        if (firstOnly)
                        {
                            conceptos = txtListaConceptos.Lines[i];
                            firstOnly = false;
                        }
                        else
                        {
                            if (txtListaConceptos.Lines[i].Trim() != "")
                                conceptos = conceptos + ";" + txtListaConceptos.Lines[i];
                        }
                    }


                    //foreach (DataGridViewRow r in this.dgvConceptos.Rows)
                    //{
                    //    if (Convert.ToBoolean(r.Cells[Sel.Name].Value) == true)
                    //    {

                    //        if (firstOnly)
                    //        {
                    //            conceptos = r.Cells[Num.Name].Value.ToString();
                    //            firstOnly = false;
                    //        }
                    //        else
                    //        {
                    //            conceptos = conceptos + ";" + r.Cells[Num.Name].Value.ToString();
                    //        }

                    //    } // if (Convert.ToBoolean(r.Cells[Sel.Name].Value) == true)

                    //} // foreach(DataGridViewRow r in this.dgvConceptos.Rows)
                    #endregion

                    Cursor.Current = Cursors.WaitCursor;
                                        
                    dgvFacturas.AutoGenerateColumns = false;
                    dgvFacturas.DataSource = 
                        inffact.CargadgvFactura("fact.fo", Convert.ToDateTime(txtFFACTURADES.Value), Convert.ToDateTime(txtFFACTURAHAS.Value),
                        usarFechaFactura, cupsree, cnifdnic, conceptos, tipoNegocio, empresas, tiposFactura, cfactura);

                    //if (dgvFacturas.DataSource.count == 0)
                    //{
                    //    MessageBox.Show("La consulta que ha realizado no devuelve ningún resultado.",
                    //"La consulta no devuelte datos.",
                    //MessageBoxButtons.OK,
                    //MessageBoxIcon.Information);
                    //}

                    Cursor.Current = Cursors.Default;

                    
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la búsqueda de facturas",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        private Boolean HayConceptosSeleccionados()
        {
           
            return txtListaConceptos.Lines.Length > 0;
        }

        private void FrmFacturasOperaciones_Load(object sender, EventArgs e)
        {

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
            this.CargadgvFacturas();
            Cursor = Cursors.Default;
        }

        private void toolTipNIF_Popup(object sender, PopupEventArgs e)
        {

        }

        private void dgvFacturas_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            DateTime ahora = new DateTime();
            Boolean firstOnly = true;
            int selectedIndex;
            Object selectedItem;
            int c = 0;

            Microsoft.Office.Interop.Excel.Application xlexcel;

            try
            {

                
                selectedItem = cmbTipoInforme.SelectedItem;
                selectedIndex = cmbTipoInforme.SelectedIndex;

                ahora = DateTime.Now;
                Cursor = Cursors.WaitCursor;
                

                switch (selectedIndex)
                {
                    #region Informe Normal
                    case 0:
                        xlexcel = new Excel.Application();
                        copyAlltoClipboard();
                        
                        Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                        Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                        object misValue = System.Reflection.Missing.Value;
                        //
                        xlexcel.Visible = true;
                        xlWorkBook = xlexcel.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        c = 2;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "EMPRESA"; c++;                        
                        xlexcel.ActiveSheet.Cells(1, c).Value = "CNIFDNIC"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "DAPERSOC"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "CUPSREE"; c++;                        
                        xlexcel.ActiveSheet.Cells(1, c).Value = "CFACTURA"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "FFACTURA"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "FFACTDES"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "FFACTHAS"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "VCUOVAFA"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "IVA"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "IIMPUES2"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "IMP_CAV"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "IFACTURA"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "CREFEREN"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "SECFACTU"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "TESTFACT"; c++;
                        xlexcel.ActiveSheet.Cells(1, c).Value = "TFACTURA"; 

                        for (int i = 1; i < 10; i++)
                        {
                            c++;
                            xlexcel.ActiveSheet.Cells(1, c).Value = "TCONFAC" + i;
                            c++;
                            xlexcel.ActiveSheet.Cells(1, c).Value = "ICONFAC" + i;
                        }

                        for (int i = 10; i <= 20; i++)
                        {
                            c++;
                            xlexcel.ActiveSheet.Cells(1, c).Value = "TCONFA" + i;
                            c++;
                            xlexcel.ActiveSheet.Cells(1, c).Value = "ICONFA" + i;
                        }

                        for (int i = 1; i <= c; i++)
                        {
                            xlexcel.ActiveSheet.Cells(1, i).Font.Bold = true;
                        }

                        Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[2, 1];
                        CR.Select();
                        xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                        break;
                    #endregion
                    case 1:

                        selectedIndex = cmbTipoFecha.SelectedIndex;
                        selectedItem = cmbTipoFecha.SelectedItem;
                        

                        if (!this.HayConceptosSeleccionados() && !otrosConceptos)
                        {
                            MessageBox.Show("Es necesario seleccionar algún concepto para sacar este informe.",
                              "Error en la selección de parámetros",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Information);
                        }else
                        {
                            EndesaBusiness.facturacion.InformeXConcepto inf =
                                new EndesaBusiness.facturacion.InformeXConcepto();

                            inf.lf = inffact.lf;
                            inf.fechaFactura = !(selectedIndex == 0);
                            inf.conceptos = conceptos;
                            inf.mainQuery = mainQuery;
                            inf.PintarExcel();
                            otrosConceptos = false;
                        }

                        break;    

                    case 2:
                        selectedIndex = cmbTipoFecha.SelectedIndex;
                        selectedItem = cmbTipoFecha.SelectedItem;

                        EndesaBusiness.facturacion.RetrasoFacturacion r =
                            new EndesaBusiness.facturacion.RetrasoFacturacion();
                        r.fechaFactura = !(selectedIndex == 0);
                        r.fd = txtFFACTURADES.Value;
                        r.fh = txtFFACTURAHAS.Value;
                        if (this.txtCUPSREE.Text != "")
                        {
                            r.cupsREE = this.txtCUPSREE.Text;
                        }
                        if (this.txtCNIFDNIC.Text != "")
                        {
                            r.cnifdnic = this.txtCNIFDNIC.Text;
                        }

                        if (listBoxLinea.SelectedItems.Count > 0)
                        {

                            selectedIndex = listBoxLinea.SelectedIndex;
                            selectedItem = listBoxLinea.SelectedItem;

                            foreach (Object values in listBoxLinea.SelectedItems)
                            {
                                if (firstOnly)
                                {

                                    r.tipoNegocio = values.ToString().Substring(0, 1);
                                    firstOnly = false;
                                }
                                else
                                {
                                    r.tipoNegocio = r.tipoNegocio + ";" + values.ToString().Substring(0, 1);
                                }
                            }

                            firstOnly = true;
                        }

                        if (listBoxEmpresas.SelectedItems.Count > 0)
                        {
                            lista_empresas.Clear();

                            foreach (Object values in listBoxEmpresas.SelectedItems)
                            {

                                lista_empresas.Add(values.ToString());

                                if (firstOnly)
                                {
                                    r.empresas = values.ToString();
                                    firstOnly = false;
                                }
                                else
                                {
                                    r.empresas = r.empresas + ";" + values.ToString();
                                }
                            }

                            firstOnly = true;

                        }

                        if (listBoxTiposFactura.SelectedItems.Count > 0)
                        {
                            foreach (Object values in listBoxTiposFactura.SelectedItems)
                            {
                                if (firstOnly)
                                {

                                    r.tipoFactura = values.ToString();
                                    firstOnly = false;
                                }
                                else
                                {
                                    r.tipoFactura = r.tipoFactura + ";" + values.ToString();
                                }
                            }

                            firstOnly = true;
                        }
                        r.CargaDatos();
                        r.PintarExcel();
                        break;
                } // switch (selectedIndex)

            
            
            }catch(Exception ee)
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

        private void dgvConceptos_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void cmdSearchConcept_Click(object sender, EventArgs e)
        {
            this.CargadgvConceptos();
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
        
        private void CambiaFecha()
        {
            anio = DateTime.Now.Year;
            mes = (DateTime.Now.Month);
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            fh = new DateTime(anio, mes - 1, dias_del_mes);
            fd = new DateTime(fh.Year, fh.Month, 1);

            txtFFACTURADES.Value = fd;
            txtFFACTURAHAS.Value = fh;
        }

        private void cmbTipoFecha_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dgvConceptos_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void cmbTipoInforme_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmdListCups_Click(object sender, EventArgs e)
        {
            
            forms.FrmListBox f = new forms.FrmListBox();          

            f.Text = "Introduzca lista de CUPS13 o CUPS20";
            f.ShowDialog(this);

            if (f.txtLista != null)
            {                
                txtListaCUPS.Text = f.txtLista.Text;
            }
                

            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool firstOnly = true;            
            
            forms.FrmListBox f = new forms.FrmListBox();
            f.Text = "Introduzca lista de NIFS";

           


            f.ShowDialog(this);

            if (f.txtLista != null)
            {
                //lista = f.txtLista.Text.Split()
                for (int i = 0; i < f.txtLista.Lines.Length; i++)
                {
                    if (firstOnly)
                    {
                        
                        listanifs = f.txtLista.Lines[i].Trim();
                        firstOnly = false;
                    }
                    else
                    {
                        if (f.txtLista.Lines[i].Trim() != "")
                            listanifs = listanifs + "|" + f.txtLista.Lines[i].Trim();
                    }
                }
            }
            else
            {
               
                listanifs = null;
            }
                


        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void cmdListConcept_Click(object sender, EventArgs e)
        {
         

            forms.FrmListBox f = new forms.FrmListBox();
            f.Text = "Introduzca lista de Conceptos 'tconfac' en sólo en número";
         

            f.ShowDialog(this);

            if (f.txtLista.Text != "")
            {
                txtListaConceptos.Text = f.txtLista.Text;
            }
               

        }

        private void cmdListNIFS_Click(object sender, EventArgs e)
        {
            forms.FrmListBox f = new forms.FrmListBox();
            f.Text = "Introduzca lista de NIF´s";
            f.ShowDialog(this);

            if (f.txtLista != null)
            {
                txtListaNIFS.Text = f.txtLista.Text;
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            
            int j = 0;
            int selectedCellCount = dgvConceptos.GetCellCount(DataGridViewElementStates.Selected);

            for (int i = 0;i < selectedCellCount; i++)
            {

                j = dgvConceptos.SelectedCells[i].RowIndex;
                if (txtListaConceptos.Text == "")
                {
                    txtListaConceptos.Text = Convert.ToString(dgvConceptos.Rows[j].Cells[1].Value);
                    
                }else
                {
                    
                    txtListaConceptos.Text =
                        txtListaConceptos.Text + System.Environment.NewLine +
                        dgvConceptos.Rows[j].Cells[0].Value;
                }
               
                    
            }
        }

        private void btnRemoveConcepts_Click(object sender, EventArgs e)
        {
            txtListaConceptos.Text = "";
        }

        private void parámetrosGlobalesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.Frm_fo_param f = new Frm_fo_param();
            //f.Show();
        }

        private void FrmFacturasOperaciones_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmFacturasOperaciones" ,"N/A");
        }
    }
}
