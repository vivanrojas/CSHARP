using GestionOperaciones.forms.medida;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmTunelCuadroMando : Form
    {
        EndesaBusiness.utilidades.Param p = new EndesaBusiness.utilidades.Param("tunel_param", EndesaBusiness.servidores.MySQLDB.Esquemas.FAC);
        EndesaBusiness.facturacion.TunelCuadroMando t;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        List<EndesaEntity.TunelContrato> lista;
        public FrmTunelCuadroMando()
        {
            usage.Start("Facturación", "FrmTunelCuadroMando" ,"N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            LoadData();
            LoadDataDetail();
            Filters();
        }

        private void LoadData()
        {
            

            t = new EndesaBusiness.facturacion.TunelCuadroMando(txt_fecha_consumo_desde.Value, txt_fecha_consumo_hasta.Value, txt_cliente.Text);

            lista = t.dic_contratos.Select(z => z.Value).ToList();

            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;
            lbl_total_contratos.Text = string.Format("Contratos: {0:#,##0}", lista.Count());

            //foreach (DataGridViewRow row in dgv.Rows)
            //{
            //    if (row.Cells[7].Value.ToString() == "No")
            //        row.DefaultCellStyle.BackColor = Color.Yellow;
            //    else
            //        row.DefaultCellStyle.BackColor = Color.Green;
            //}
                
        }

        private void LoadDataDetail()
        {
            dgvd.AutoGenerateColumns = false;
            double total_energia = 0;
            
            if(t.dic_contratos.Count > 0)
            {
                dgvd.DataSource = t.dic_contratos.FirstOrDefault().Value.lista;
                foreach(EndesaEntity.Tunel p in t.dic_contratos.FirstOrDefault().Value.lista)
                {
                    total_energia += (p.total_energia / 1000000);
                }               
                lbl_total_energia.Text = string.Format("Total Energía (GWh): {0:#,##0.00}", total_energia);
            }
            


        }

        private void Filters()
        {
            List<EndesaEntity.TunelContrato> lista_temp =
                new List<EndesaEntity.TunelContrato>();

            List<int> lista_lotes = new List<int>();

            if (chk_aplica_tunel.Checked)
                lista = lista.Where(z => z.aplica_tunel).ToList();

            if (chk_formula_antigua.Checked)
                lista = lista.Where(z => z.formula_antigua).ToList();

            if (chk_ltps.Checked)
                lista = lista.Where(z => z.ltps).ToList();

            if (chk_baja_tension.Checked)
                lista = lista.Where(z => z.baja_tension).ToList();

            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;
            lbl_total_contratos.Text = string.Format("Contratos: {0:#,##0}", lista.Count());
        }

        private void FrmTunelCuadroMando_Load(object sender, EventArgs e)
        {
            DateTime mesAnterior = DateTime.Now;
            //DateTime mesAnterior = DateTime.Now;
            int anio = mesAnterior.Year -1;
            int mes = 12;
            int dias_del_mes = DateTime.DaysInMonth(anio, mes);

            DateTime fh = new DateTime(anio, 12, dias_del_mes);
            DateTime fd = new DateTime(fh.Year, 1, 1);

            txt_fecha_consumo_desde.Value = fd;
            txt_fecha_consumo_hasta.Value = fh;

            

            LoadData();
            LoadDataDetail();
        }

        private void importarContratosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            DialogResult result = DialogResult.Yes;

            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Seleccione archivos Excel de Tunel";
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
               
                    Cursor.Current = Cursors.WaitCursor;

                    foreach (string fileName in d.FileNames)
                    {
                        t.CargaExcel(fileName, false);
                    }
                    Cursor.Current = Cursors.Default;

                    MessageBox.Show("Proceso Finalizado.",                        
                  "Importación ficheros Excel Contratos Tunel",
                  MessageBoxButtons.OK, MessageBoxIcon.Information);

                LoadData();
                LoadDataDetail();


            }
        }

        private void dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            double total_energia = 0;
            if (e.RowIndex >= 0)
            {

                btnEdit.Enabled = true;
                DataGridViewCell cliente = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[0];

                DataGridViewCell start_date = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[1];

                DataGridViewCell end_date = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[2];



                dgvd.AutoGenerateColumns = false;
                EndesaEntity.TunelContrato o;
                if (t.dic_contratos.TryGetValue(cliente.Value.ToString() 
                    + Convert.ToDateTime(start_date.Value).ToString("yyyyMMdd") 
                    + Convert.ToDateTime(end_date.Value).ToString("yyyyMMdd"), out o))
                {
                    dgvd.DataSource = o.lista;
                    foreach (EndesaEntity.Tunel p in o.lista)
                    {
                        total_energia += p.total_energia;
                    }
                    lbl_total_energia.Text = string.Format("Total Energía: {0:#,##0.00}", total_energia);
                }
                    
                else
                    dgvd.DataSource = null;
                // dgvd.DataSource = contratos.cgd.dic.FirstOrDefault(z => z.Key == cups20.Value.ToString()).Value.OrderByDescending(z => z.fecha_inicio).ToList();                

            }
        }

        private void dgv_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            t.InformeExcel();
        }

        private void contratosAExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            t.InformeExcel();
        }

        private void FrmTunelCuadroMando_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmTunelCuadroMando" ,"N/A");
        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización Tunel";
            p.tabla = "tunel_param";
            p.esquemaString = "FAC";
            p.Show();
        }

        private void dgvd_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            btnEdit.Enabled = false;
        }

        private void btnEdit_d_Click(object sender, EventArgs e)
        {

        }

        private void Edit()
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            int c = 0;

            FrmTunelCuadroMando_Edit f = new FrmTunelCuadroMando_Edit();
            f.Text = "Editar contrato";
            f.tunel = t;

            f.txt_cliente.Enabled = false;
            f.txt_fecha_inicio_tunel.Enabled = false;
            f.txt_fecha_fin_tunel.Enabled = false;

            row = dgv.Rows[fila];

            if (row.Cells[c].Value != null)
                f.txt_cliente.Text = row.Cells[c].Value.ToString(); c++;            

            if (row.Cells[c].Value != null)
                f.txt_fecha_inicio_tunel.Value = Convert.ToDateTime(row.Cells[c].Value); c++;

            if (row.Cells[c].Value != null)
                f.txt_fecha_fin_tunel.Value = Convert.ToDateTime(row.Cells[c].Value); c++;

            if (row.Cells[12].Value != null)
                f.chk_formula_antigua.Checked = Convert.ToBoolean(row.Cells[12].Value);

            f.ShowDialog();
            LoadData();
            LoadDataDetail();
            Filters();

        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            Edit();
        }

        private void importarConsumosBTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = DialogResult.Yes;

            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Seleccione archivos Excel consumos BT";
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                Cursor.Current = Cursors.WaitCursor;

                foreach (string fileName in d.FileNames)
                {
                    t.CargaExcelConsumos(fileName, false);
                }
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Proceso Finalizado.",
                    "Importación ficheros Excel consumos BT",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                LoadData();
                LoadDataDetail();
            }
        }
    }
}
