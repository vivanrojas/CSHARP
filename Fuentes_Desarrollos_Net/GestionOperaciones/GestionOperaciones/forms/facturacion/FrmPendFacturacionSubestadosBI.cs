using EndesaBusiness.factoring;
using EndesaEntity.factoring;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmPendFacturacionSubestadosBI : Form
    {
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        EndesaBusiness.facturacion.redshift.Pendiente_Subestados subestados;
        List<EndesaEntity.medida.Pendiente> lista;
        
        public FrmPendFacturacionSubestadosBI()
        {
            usage.Start("Facturación", "FrmPendFacturacionSubestadosBI", "Subestadps");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmPendFacturacionSubestadosBI_Load(object sender, EventArgs e)
        {
            LoadData();
            Filters();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmPendFacturacionSubestadosBi_Edit f = new FrmPendFacturacionSubestadosBi_Edit();
            f.Text = "Nuevo registro";
            f.pendiente = subestados;
            f.ShowDialog();
            LoadData();
            Filters();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            int c = 0;

            forms.facturacion.FrmPendFacturacionSubestadosBi_Edit f = new FrmPendFacturacionSubestadosBi_Edit();
            f.Text = "Editar registro";

            f.pendiente = subestados;
            row = dgv.Rows[fila];

            if (row.Cells[c].Value != null)
                f.txt_cod_subestado.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_subestado_descripcion.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.cmb_area_responsable.Text = row.Cells[c].Value.ToString(); c++;

            f.ShowDialog();

            LoadData();
            Filters();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            int c = 0;

            row = dgv.Rows[fila];

            if (row.Cells[c].Value != null)
                subestados.cod_subestado = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                subestados.descripcion_subestado = row.Cells[c].Value.ToString(); c++;


            subestados.Del();

            LoadData();
            Filters();
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {

        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            LoadData();
            Filters();
        }

        private void FrmPendFacturacionSubestadosBI_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.Start("Facturación", "FrmPendFacturacionSubestadosBI", "Subestadps");
        }

        private void LoadData() 
        {
            subestados = new EndesaBusiness.facturacion.redshift.Pendiente_Subestados();
            lista = new List<EndesaEntity.medida.Pendiente>();

            foreach (KeyValuePair<string, EndesaEntity.medida.Pendiente> p in subestados.dic)
                lista.Add(p.Value);



            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;
            lbl_total_cups.Text = string.Format("Total Registros: {0:#,##0}", lista.Count());
        }

        private void Filters()
        {
            dgv.AutoGenerateColumns = false;

            if (txt_nif.Text != string.Empty)
                lista = lista.Where(z => z.cod_subestado.Contains(txt_nif.Text.ToUpper())).ToList();

            dgv.DataSource = lista;
            lbl_total_cups.Text = string.Format("Total Registros: {0:#,##0}", lista.Count());

        }
    }
}
