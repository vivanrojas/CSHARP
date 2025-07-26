using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.medida
{
   

    public partial class FrmAdifInventario : Form
    {
        List<EndesaEntity.ListaLotes> ll = new List<EndesaEntity.ListaLotes>();
        
        EndesaBusiness.adif.InventarioFunciones inventario;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        List<EndesaEntity.medida.AdifInventario> lista;

        public FrmAdifInventario()
        {
            int anio;
            int mes;
            int dias_del_mes;
            DateTime fh = new DateTime();
            DateTime fd = new DateTime();
            DateTime mesAnterior = new DateTime();


            usage.Start("Medida", "FrmAdifInventario", "N/A");
            InitializeComponent();

            mesAnterior = DateTime.Now.AddMonths(-1);
            anio = mesAnterior.Year;
            mes = mesAnterior.Month;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            fh = new DateTime(anio, mes, dias_del_mes);
            fd = new DateTime(fh.Year, fh.Month, 1);

            txtFD.Value = fd;
            txtFH.Value = fh;

            InitListBox(fd, fh);
        }

        private void InitListBox(DateTime fd, DateTime fh)
        {

            EndesaBusiness.adif.AdifLotes lotes = new EndesaBusiness.adif.AdifLotes(fd, fh);

            try
            {
                for (int i = listBoxLotes.Items.Count - 1; i >= 0; i--)
                {
                    listBoxLotes.Items.RemoveAt(i);
                }



                for(int i = 0; i < lotes.lista_lotes.Count; i++)                
                    listBoxLotes.Items.Add(lotes.lista_lotes[i]);
                
                
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "InitListBox",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            }
            
        }

        private void LoadData()
        {
            List<string> lista_lotes = new List<string>();

            if (listBoxLotes.SelectedItems.Count > 0)
                foreach (Object values in listBoxLotes.SelectedItems)
                    lista_lotes.Add(values.ToString());

            inventario = new EndesaBusiness.adif.InventarioFunciones(Convert.ToDateTime(txtFD.Text), Convert.ToDateTime(txtFH.Text), null, lista_lotes);

            this.CreaListaLotes();

            lista = inventario.dic_inventario.Select(z => z.Value).ToList();

            dgvInventario.AutoGenerateColumns = false;            
            dgvInventario.DataSource = lista;
            lbl_total_cups.Text = string.Format("Total Registros: {0:#,##0}", lista.Count());

        }

        private void Filters()
        {
            List<EndesaEntity.medida.AdifInventario> lista_temp =
                new List<EndesaEntity.medida.AdifInventario>();

            List<int> lista_lotes = new List<int>();

            if (listBoxLotes.SelectedItems.Count > 0)
                foreach (Object values in listBoxLotes.SelectedItems) 
                    lista_lotes.Add(Convert.ToInt32(values));


            if (listBoxLotes.SelectedItems.Count > 0)
            {
                foreach (EndesaEntity.medida.AdifInventario p in lista)
                {
                    foreach (int pp in lista_lotes)
                    {
                        if (p.lote == pp)
                            lista_temp.Add(p);
                    }

                }

                lista = lista_temp;
            }
               

            if (txtcups20.Text != string.Empty)
                lista = lista.Where(z => z.cups20.Contains(txtcups20.Text.ToUpper())).ToList();

            if(chk_cierres_energia.Checked)
                lista = lista.Where(z => z.cierres_energia).ToList();

            if (chk_devolucion_energia.Checked)
                lista = lista.Where(z => z.devolucion_de_energia).ToList();

            if (chk_medida_en_baja.Checked)
                lista = lista.Where(z => z.medida_en_baja).ToList();
            if(chk_perdidas.Checked)
                lista = lista.Where(z => z.tiene_perdidas).ToList();  

            dgvInventario.AutoGenerateColumns = false;
            dgvInventario.DataSource = lista;
            lbl_total_cups.Text = string.Format("Total Registros: {0:#,##0}", lista.Count());

        }


        private void FrmAdifInventario_Load(object sender, EventArgs e)
        {
            this.LoadData();
        }

        private void dgvInventario_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex >= 0)
            //{
            //    DataGridViewCell cups20 = (DataGridViewCell)
            //        dgvInventario.Rows[e.RowIndex].Cells[0];

            //    invf.GetInventario(cups20.Value.ToString());                
            //    this.txt_e_cups20.Text = invf.cups20;
            //    this.txt_e_lote.Text = invf.lote.ToString();
            //    this.txt_e_fd.Value = invf.vigencia_desde;
            //    this.txt_e_fh.Value = invf.vigencia_hasta;
            //    this.txt_e_tarifa.Text = invf.tarifa;
            //    this.txt_e_dustribuidora.Text = invf.distribuidora;
            //    this.txt_e_tension.Text = invf.tension.ToString();
            //    this.txt_e_zona.Text = invf.zona;
            //    this.txt_e_codigo.Text = invf.codigo;
            //    this.txt_e_dustribuidora.Text = invf.distribuidora;
            //    this.txt_e_comentario.Text = invf.comentarios;
            //    this.chk_cierres_energia.Checked = invf.cierres_energia;
            //    this.chk_devolucion_energia.Checked = invf.devolucion_de_energia;
            //    this.chk_medida_en_baja.Checked = invf.medida_en_baja;

            //}
                
        }

        private void CreaListaLotes()
        {
            int f;
            int c;
            int total_cups = 0;

            Dictionary<int, EndesaEntity.ListaLotes> dic_lotes = new Dictionary<int, EndesaEntity.ListaLotes>();




            foreach (KeyValuePair<string, EndesaEntity.medida.AdifInventario> p in inventario.dic_inventario)
            {

                EndesaEntity.ListaLotes lotes;
                if (!dic_lotes.TryGetValue(p.Value.lote, out lotes))
                {
                    EndesaEntity.ListaLotes ll = new EndesaEntity.ListaLotes();
                    ll.lote = p.Value.lote;
                    ll.num_cups = 1;
                    dic_lotes.Add(p.Value.lote, ll);
                }
                else
                {
                    lotes.num_cups = lotes.num_cups + 1;
                }

            }

            this.dvgLotes.Rows.Clear();
            dvgLotes.RowCount = dic_lotes.Count() + 1;
            f = 0;
            c = 0;
            foreach (KeyValuePair<int, EndesaEntity.ListaLotes> p in dic_lotes)
            {
                c = 0;
                dvgLotes.Rows[f].Cells[c].Value = p.Key; c++;
                dvgLotes.Rows[f].Cells[c].Value = p.Value.num_cups; c++;
                total_cups += p.Value.num_cups;
                f++;
            }
            c = 0;
            dvgLotes.Rows[f].Cells[c].Value = "Total:"; c++;
            dvgLotes.Rows[f].Cells[c].Value = total_cups;

        }

        private void txt_e_comentario_TextChanged(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            LoadData();
            Filters();
        }

        

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del archivo inventario ADIF";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                this.ExportExcel(save.FileName);
                MessageBox.Show("Informe terminado.",
                  "Exportación a Excel",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Information);

                System.Diagnostics.Process.Start(save.FileName);

            }
        }

        private void FrmAdifInventario_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Medida", "FrmAdifInventario", "N/A");
        }

        private void ExportExcel(string fichero)
        {
            int f = 1;
            int c = 1;

            FileInfo fileInfo = new FileInfo(fichero);

            if (fileInfo.Exists)
                fileInfo.Delete();


            ExcelPackage excelPackage;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            excelPackage = new ExcelPackage(fileInfo);
            //ExcelPackage excelPackage = new ExcelPackage();
            var workSheet = excelPackage.Workbook.Worksheets.Add("Inventario");



            var headerCells = workSheet.Cells[1, 1, 1, 27];
            var headerFont = headerCells.Style.Font;


            
            headerFont.Bold = true;
            

            workSheet.Cells[f, c].Value = "CUPS20"; c++;            
            workSheet.Cells[f, c].Value = "LOTE"; c++;
            workSheet.Cells[f, c].Value = "FECHA DESDE"; c++;
            workSheet.Cells[f, c].Value = "FECHA HASTA"; c++;
            workSheet.Cells[f, c].Value = "ZONA"; c++;
            workSheet.Cells[f, c].Value = "CÓDIGO"; c++;
            workSheet.Cells[f, c].Value = "NOMBRE PUNTO SUMINISTRO"; c++;
            workSheet.Cells[f, c].Value = "DISTRIBUIDORA"; c++;
            workSheet.Cells[f, c].Value = "COMENTARIOS"; c++;
            workSheet.Cells[f, c].Value = "TARIFA"; c++;
            workSheet.Cells[f, c].Value = "TENSIÓN"; c++;
            workSheet.Cells[f, c].Value = "P1"; c++;
            workSheet.Cells[f, c].Value = "P2"; c++;
            workSheet.Cells[f, c].Value = "P3"; c++;
            workSheet.Cells[f, c].Value = "P4"; c++;
            workSheet.Cells[f, c].Value = "P5"; c++;
            workSheet.Cells[f, c].Value = "P6"; c++;
            workSheet.Cells[f, c].Value = "Medida en baja"; c++;
            workSheet.Cells[f, c].Value = "Devolución de energía"; c++;
            workSheet.Cells[f, c].Value = "Cierres energía"; c++;
            workSheet.Cells[f, c].Value = "Provincia"; c++;
            workSheet.Cells[f, c].Value = "Comunidad Autónoma"; c++;
            workSheet.Cells[f, c].Value = "Sistema tracción"; c++;
            workSheet.Cells[f, c].Value = "Grupo"; c++;
            workSheet.Cells[f, c].Value = "Valor KVA´s"; c++;
            workSheet.Cells[f, c].Value = "Pérdidas"; c++;
            workSheet.Cells[f, c].Value = "Nº Multipunto principales"; c++;


            // foreach (KeyValuePair<string, EndesaEntity.medida.AdifInformeMedida> p in !chkDiff.Checked ? med.dic_informe : med.dic_informe.Where(kvp => kvp.Value.dif_sce_adif_a != 0))
            //foreach (KeyValuePair<string, EndesaEntity.medida.AdifInventario> p in inventario.dic_inventario)
            foreach (EndesaEntity.medida.AdifInventario p in lista)
            {
                f++;
                c = 1;
                workSheet.Cells[f, c].Value = p.cups20; c++;
                workSheet.Cells[f, c].Value = p.lote; c++;                
                workSheet.Cells[f, c].Value = p.vigencia_desde; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.vigencia_hasta; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.zona; c++;
                workSheet.Cells[f, c].Value = p.codigo; c++;
                workSheet.Cells[f, c].Value = p.nombre_punto_suministro; c++;
                workSheet.Cells[f, c].Value = p.distribuidora; c++;
                workSheet.Cells[f, c].Value = p.comentarios; c++;
                workSheet.Cells[f, c].Value = p.tarifa; c++;
                workSheet.Cells[f, c].Value = p.tension; c++;

                for(int i = 0; i < 6; i++)
                {
                    workSheet.Cells[f, c].Value = p.p[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##00"; 
                    c++;
                }

                if (p.medida_en_baja)
                    workSheet.Cells[f, c].Value = "Sí";
                else
                    workSheet.Cells[f, c].Value = "No";

                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                if (p.devolucion_de_energia)
                    workSheet.Cells[f, c].Value = "Sí";
                else
                    workSheet.Cells[f, c].Value = "No";

                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                if (p.cierres_energia)
                    workSheet.Cells[f, c].Value = "Sí";
                else
                    workSheet.Cells[f, c].Value = "No";

                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                if(p.provincia != null)
                    workSheet.Cells[f, c].Value = p.provincia;

                c++;

                if (p.comunidad_autonoma != null)
                    workSheet.Cells[f, c].Value = p.comunidad_autonoma;

                c++;

                if (p.sitema_traccion != null)
                    workSheet.Cells[f, c].Value = p.sitema_traccion;

                c++;

                if (p.grupo != null)
                    workSheet.Cells[f, c].Value = p.grupo;

                c++;

                if (p.valor_kvas > 0)
                    workSheet.Cells[f, c].Value = p.valor_kvas;

                c++;

                if (p.perdidas > 0)
                    workSheet.Cells[f, c].Value = p.perdidas;

                c++;

                if (p.multipunto_num_principales > 0)
                    workSheet.Cells[f, c].Value = p.multipunto_num_principales;

                c++;

            }


            var allCells = workSheet.Cells[1, 1, f, c];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:AA1"].AutoFilter = true;
            allCells.AutoFitColumns();
            excelPackage.Save();


        }

        private void btn_ficheros_Click(object sender, EventArgs e)
        {

        }

        private void chk_devolucion_energia_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void dgvInventario_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex >= 0)
            //{
            //    DataGridViewCell cups20 = (DataGridViewCell)
            //        dgvInventario.Rows[e.RowIndex].Cells[0];

            //    invf.GetInventario(cups20.Value.ToString());
            //    this.txt_e_cups20.Text = invf.cups20;
            //    this.txt_e_lote.Text = invf.lote.ToString();
            //    this.txt_e_fd.Value = invf.vigencia_desde;
            //    this.txt_e_fh.Value = invf.vigencia_hasta;
            //    this.txt_e_tarifa.Text = invf.tarifa;
            //    this.txt_e_dustribuidora.Text = invf.distribuidora;
            //    this.txt_e_tension.Text = invf.tension.ToString();
            //    this.txt_e_zona.Text = invf.zona;
            //    this.txt_e_codigo.Text = invf.codigo;
            //    this.txt_e_dustribuidora.Text = invf.distribuidora;
            //    this.txt_e_comentario.Text = invf.comentarios;
            //    this.chk_cierres_energia.Checked = invf.cierres_energia;
            //    this.chk_devolucion_energia.Checked = invf.devolucion_de_energia;
            //    this.chk_medida_en_baja.Checked = invf.medida_en_baja;

            //}
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            forms.medida.FrmAdifIventario_Edit f = new FrmAdifIventario_Edit();
            f.Text = "Nuevo registro";
            f.inventario = inventario;
            f.ShowDialog();
            LoadData();
            Filters();
        }

        

        private void btnEdit_Click_1(object sender, EventArgs e)
        {
            Edit();
        }

        private void Edit()
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgvInventario.CurrentRow.Index;
            int c = 0;

            FrmAdifIventario_Edit f = new FrmAdifIventario_Edit();
            f.Text = "Editar cierre";
            f.inventario = inventario;
            f.txt_cups20.Enabled = false;

            row = dgvInventario.Rows[fila];

            if (row.Cells[c].Value != null)
                f.txt_cups20.Text = row.Cells[c].Value.ToString(); c++;
            c++;

            if (row.Cells[c].Value != null)
                f.txt_fd.Value = Convert.ToDateTime(row.Cells[c].Value); c++;

            if (row.Cells[c].Value != null)
                f.txt_fh.Value = Convert.ToDateTime(row.Cells[c].Value); c++;

            inventario.GetRegistroCUPS20(f.txt_cups20.Text);

            f.txt_lote.Text = inventario.lote.ToString();
            f.txt_distribuidora.Text = inventario.distribuidora;
            f.txt_zona.Text = inventario.zona;
            f.txt_codigo.Text = inventario.codigo;
            f.txt_nombre_suministro.Text = inventario.nombre_punto_suministro;
            f.txt_distribuidora.Text = inventario.distribuidora;
            f.txt_tarifa.Text = inventario.tarifa;            
            f.txt_p1.Text = inventario.p[0].ToString();
            f.txt_p2.Text = inventario.p[1].ToString();
            f.txt_p3.Text = inventario.p[2].ToString();
            f.txt_p4.Text = inventario.p[3].ToString();
            f.txt_p5.Text = inventario.p[4].ToString();
            f.txt_p6.Text = inventario.p[5].ToString();
            f.txt_comentarios.Text = inventario.comentarios;
            f.chk_devolucion_energia.Checked = inventario.devolucion_de_energia;
            f.chk_medida_en_baja.Checked = inventario.medida_en_baja;
            f.cmb_provincia.Text = inventario.provincia;
            f.txt_comunidad_autonoma.Text = inventario.comunidad_autonoma;
            f.txt_sistema_traccion.Text = inventario.sitema_traccion;
            f.txt_grupo.Text = inventario.grupo;
            f.txt_valor_kvas.Text = inventario.valor_kvas.ToString();
            f.txt_perdidas.Text = inventario.perdidas.ToString();
            f.txt_multipuntos.Text = inventario.multipunto_num_principales.ToString();

            f.ShowDialog();
            LoadData();
            Filters();

        }

        private void chk_cierres_energia_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
