using EndesaBusiness.servidores;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GestionOperaciones.forms.contratacion.gas
{
    
    public partial class FrmDistribuidoras : Form
    {

        List<EndesaEntity.Table_atrgas_distribuidoras> filtro;
        EndesaBusiness.contratacion.gestionATRGas.Distribuidoras distribuidoras;
        EndesaBusiness.contratacion.gestionATRGas.Excepciones excepciones;

        String mainQuery = null;
        EndesaBusiness.utilidades.Global g = new EndesaBusiness.utilidades.Global();

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        public FrmDistribuidoras()
        {
            usage.Start("Contratación", "FrmDistribuidoras", "Distribuidoras");
            InitializeComponent();
            Cargadgv();
            //Cargamos excepciones, pendientes y activas en dgv_excepciones y finalizadas en dgv_excepciones_finalizadas
            Cargadgv_excepciones();
            
            if (dgv_excepciones.RowCount > 0)
            {
                MessageBox.Show("Atención: hay excepciones programadas para tramitaciones por email de distribuidoras ",
                "Excepciones programadas para tramitación por eMail",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
                tabControl1.SelectedTab = tabExcpeciones;
            }
        }

        private void Cargadgv_excepciones()
        {
            
            try
            {
  
                excepciones = new EndesaBusiness.contratacion.gestionATRGas.Excepciones();

                Cursor.Current = Cursors.WaitCursor;

                #region Excepciones programadas y en ejecución
                dgv_excepciones.AutoGenerateColumns = false;
                dgv_excepciones.DataSource = excepciones.lista_excepcion_tramitacion.Where(x => x.estado != "Finalizada" && x.estado != "Cancelada").ToList();
                #endregion

                #region Excepciones finalizadas
                dgv_historico_excepciones.AutoGenerateColumns = false;
                dgv_historico_excepciones.DataSource = excepciones.lista_excepcion_tramitacion.Where(x => x.estado == "Finalizada" || x.estado == "Cancelada").ToList();
                #endregion

                Cursor.Current = Cursors.Default;


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la carga de excepciones de tramitación",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        private void Cargadgv()
        {            
            DateTime begin = new DateTime();
            

            try
            {

                distribuidoras = new EndesaBusiness.contratacion.gestionATRGas.Distribuidoras(true);

                begin = DateTime.Now;                              

                Cursor.Current = Cursors.WaitCursor;                

                dgv.AutoGenerateColumns = false;

                if (txtSearch.Text != null || txtSearch.Text != "")                    
                    filtro = distribuidoras.l_distribuidoras.Where(z => z.Value.codigo.Contains(txtSearch.Text) ||
                    z.Value.nombre.Contains(txtSearch.Text) || z.Value.email.Contains(txtSearch.Text)).Select(z => z.Value).ToList();
                else
                    filtro = distribuidoras.l_distribuidoras.Values.ToList();

                dgv.DataSource = filtro;

                if (distribuidoras.l_distribuidoras.Count() == 0)
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
                "Error en la búsqueda de complementos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
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
            this.Cargadgv();
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
                
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlexcel.ActiveSheet.Cells(1, 2).Value = "Complemento";                
                xlexcel.ActiveSheet.Cells(1, 3).Value = "Descripción";
                

                 c = 5;
                for (int i = 1; i <= c; i++)
                {
                    xlexcel.ActiveSheet.Cells(1, i).Font.Bold = true;
                }

                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[2, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                        
                #endregion
                                   
            
            
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
            dgv.SelectAll();
            DataObject dataObj = dgv.GetClipboardContent();
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

            EndesaBusiness.facturacion.FestivosElectricos fes = new EndesaBusiness.facturacion.FestivosElectricos();
            GestionOperaciones.forms.FormFacFestivosElectricosParam_Edit f = new forms.FormFacFestivosElectricosParam_Edit();
            f.Text = "Añadir nuevo registro";
            f.ShowDialog(this);

            if (f.txtFecha.Value != null &&                 
                f.txtDescripcion.Text != null)
            {
                fes.fechaFestivo = f.txtFecha.Value;
                fes.descripcion = f.txtDescripcion.Text;                
                
                fes.Add();
                Cargadgv();
            }

        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            EndesaBusiness.facturacion.FestivosElectricos fes = new EndesaBusiness.facturacion.FestivosElectricos();

            try
            {
                DataGridViewRow row = new DataGridViewRow();
                int fila = dgv.CurrentRow.Index;

                forms.FormFacFestivosElectricosParam_Edit f = new forms.FormFacFestivosElectricosParam_Edit();
                f.Text = "Editar registro";
                f.txtFecha.Enabled = false;

                row = dgv.Rows[fila];
                fes.fechaFestivo = Convert.ToDateTime(row.Cells[0].Value.ToString());                
                fes.descripcion = row.Cells[1].Value.ToString();                

                f.txtFecha.Value = fes.fechaFestivo;                
                f.txtDescripcion.Text = fes.descripcion;
                

                f.ShowDialog(this);

                if (f.txtDescripcion.Text != fes.descripcion)
                    
                {

                    fes.descripcion = f.txtDescripcion.Text;
                    
                    fes.Update();
                    Cargadgv();

                }

            }catch(Exception ee)
            {
                MessageBox.Show(ee.Message,
               "Error en la edición de los datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
            
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            EndesaBusiness.facturacion.FestivosElectricos fes = new EndesaBusiness.facturacion.FestivosElectricos();
            DialogResult r;

            try
            {
                DataGridViewRow row = new DataGridViewRow();
                int fila = dgv.CurrentRow.Index;
                row = dgv.Rows[fila];
                fes.fechaFestivo = Convert.ToDateTime(row.Cells[0].Value);

                r = MessageBox.Show("¿Desea borrar el registros con Fecha " + fes.fechaFestivo + "?",
                    "Borrar registro con código " + fes.fechaFestivo,
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

               if (r == DialogResult.Yes)
                {
                    fes.Del();
                    Cargadgv();
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

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void acerdaDeDistribuidorasGASToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmDistribuidoras_Ayuda f = new FrmDistribuidoras_Ayuda();
            f.Show();
        }

        private void btnEdit_Click_1(object sender, EventArgs e)
        {

            try
            {

                DataGridViewRow row = new DataGridViewRow();
                int fila = dgv.CurrentRow.Index;

                forms.contratacion.gas.FrmDistribuidoras_Edit f = new FrmDistribuidoras_Edit();
                
                f.Text = "Editar registro";
                f.nuevo_registro = false;
                row = dgv.Rows[fila];
                f.txt_codigo.Text = row.Cells[0].Value.ToString();
                f.txt_nombre.Text = row.Cells[1].Value.ToString();
                f.txt_fecha_desde.Text = row.Cells[2].Value.ToString();
                f.txt_fecha_hasta.Text = row.Cells[3].Value.ToString();
                f.txt_email.Text = row.Cells[4].Value.ToString();

                f.ShowDialog();
                Cargadgv();
                

            }
            catch (Exception ee)

            {
                MessageBox.Show(ee.Message,
              "Error en la actualización de los datos",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }

        }

        private void btnAdd_Click_1(object sender, EventArgs e)
        {
            FrmDistribuidoras_Edit f = new FrmDistribuidoras_Edit();
            f.Text = "Nuevo registro";
            f.nuevo_registro = true;
            f.ShowDialog();
            
        }

        private void CmdExcel_Click_1(object sender, EventArgs e)
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
                ExportExcel(save.FileName);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }
        }

        private void ExportExcel(string rutaFichero)
        {

            int f = 0; // fila
            int c = 0; // columna

            FileInfo file = new FileInfo(rutaFichero);

            if (file.Exists)
                file.Delete();

            ExcelPackage excelPackage = new ExcelPackage(file);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Distribuidoras");

            var headerCells = workSheet.Cells[1, 1, 1, 5];
            var headerFont = headerCells.Style.Font;

            f++;
            c = 1;
            workSheet.Cells[f, c].Value = "Código Distribuidora"; c++;
            workSheet.Cells[f, c].Value = "Nombre Distribuidora"; c++;
            workSheet.Cells[f, c].Value = "Fecha Desde"; c++;
            workSheet.Cells[f, c].Value = "Fecha Hasta"; c++;
            workSheet.Cells[f, c].Value = "mail"; c++;

            for(int i = 0; i < filtro.Count(); i++)
            {
                f++;
                c = 1;
                workSheet.Cells[f, c].Value = filtro[i].codigo; c++;
                workSheet.Cells[f, c].Value = filtro[i].nombre; c++;
                workSheet.Cells[f, c].Value = filtro[i].fecha_desde;
                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = filtro[i].fecha_hasta;
                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = filtro[i].email; c++;
            }
            
            var allCells = workSheet.Cells[1, 1, f, c];
            allCells.AutoFitColumns();

            headerFont.Bold = true;
            allCells = workSheet.Cells[1, 1, f, 10];            
            workSheet.Cells["A1:E1"].AutoFilter = true;
            allCells.AutoFitColumns();
            workSheet.View.FreezePanes(2, 1);

            excelPackage.Save();

            MessageBox.Show("Informe terminado.",
             "Exportación a Excel",
             MessageBoxButtons.OK,
             MessageBoxIcon.Information);
                       
        }

        private void BtnDel_Click_1(object sender, EventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            int c = 0;
            row = dgv.Rows[fila];

            DialogResult result = MessageBox.Show("¿Desea borrar la distribuidora " + row.Cells[0].Value.ToString() + "?", "Borrar distribuidora",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                EndesaBusiness.contratacion.gestionATRGas.Distribuidoras d = new EndesaBusiness.contratacion.gestionATRGas.Distribuidoras();
                if (row.Cells[0].Value != null)
                    d.codigo = row.Cells[0].Value.ToString();

                d.Del();

            }
                

        }

        private void CmdSearch_Click_1(object sender, EventArgs e)
        {
            Cargadgv();
        }

        private void FrmDistribuidoras_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Contratación", "FrmDistribuidoras", "Distribuidoras");
        }

        private void btn_add_excepcion_Click(object sender, EventArgs e)
        {
            FrmExcepcion_Edit f = new FrmExcepcion_Edit();
            f.Text = "Añadir nueva excepción";
            f.nuevo_registro = true;
            f.dateTimePicker_fd.Value = DateTime.Now;
            f.dateTimePicker_fh.Value = DateTime.Now;
            //Cargar combo con los distintos nombres de distribuidoras cuyo método de tramitación sea distinto a Mail
            f.cmb_lista_distribuidoras.DataSource = excepciones.GetNombresDistribuidoras();
            
            f.ShowDialog();
            Cargadgv_excepciones();
        }

        private void btn_edit_excepcion_Click(object sender, EventArgs e)
        {
           
            try
            {
                if (dgv_excepciones.Rows.Count > 0)
                {
                    DataGridViewRow row = new DataGridViewRow();
                    int fila = dgv_excepciones.CurrentRow.Index;

                    forms.contratacion.gas.FrmExcepcion_Edit f = new FrmExcepcion_Edit();

                    f.nuevo_registro = false;

                    row = dgv_excepciones.Rows[fila];

                    f.id_excepcion = Convert.ToInt32(row.Cells[0].Value);
                    f.dateTimePicker_fd.Value = Convert.ToDateTime(row.Cells[3].Value);
                    f.dateTimePicker_fh.Value = Convert.ToDateTime(row.Cells[4].Value);
                    //Cargar combo con los distintos nombres de distribuidoras cuyo método de tramitación sea distinto a Mail
                    f.cmb_lista_distribuidoras.DataSource = excepciones.GetNombresDistribuidoras();
                    f.cmb_lista_distribuidoras.Text = row.Cells[1].Value.ToString();

                    f.Text = "Modificar datos excepción [ID " + f.id_excepcion + "]";
                    f.ShowDialog();
                    Cargadgv_excepciones();
                }

            }
            catch (Exception ee)

            {
                MessageBox.Show(ee.Message,
              "Error en la actualización de los datos de la excepción",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }
        }

        private void btn_del_excepcion_Click(object sender, EventArgs e)
        {
            if (dgv_excepciones.Rows.Count > 0)
            {
                DataGridViewRow row = new DataGridViewRow();
                int fila = dgv_excepciones.CurrentRow.Index;
                row = dgv_excepciones.Rows[fila];
                Int32 id = Convert.ToInt32(row.Cells[0].Value);

                DialogResult result = MessageBox.Show("¿Está seguro que desea cancelar la excepción seleccionada con ID " + id + "?", "Cancelar excepción",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    excepciones.ActualizaEstadoExcepcion(id, "Cancelada");
                    MessageBox.Show("La excepción en tramitación con ID " + id + " se ha cancelado correctamente",
                                    "Excepción cancelada",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                    Cargadgv_excepciones();
                }
            }
        }
    }
}
