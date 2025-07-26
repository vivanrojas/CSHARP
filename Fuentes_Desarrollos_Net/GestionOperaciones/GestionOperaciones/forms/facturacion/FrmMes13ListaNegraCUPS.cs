using System;
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
    public partial class FrmMes13ListaNegraCUPS : Form
    {
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        EndesaBusiness.factoring.Lista_Negra_Cups lista_negra;
        List<EndesaEntity.factoring.ListaNegra_CUPS> lista;
        public FrmMes13ListaNegraCUPS()
        {
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {

            LoadData();
            Filters();
        }

        private void FrmMes13ListaNegraCUPS_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void LoadData()
        {
            lista_negra = new EndesaBusiness.factoring.Lista_Negra_Cups();

            lista = new List<EndesaEntity.factoring.ListaNegra_CUPS>();

            foreach (KeyValuePair<string, EndesaEntity.factoring.ListaNegra_CUPS> p in lista_negra.dic)
                lista.Add(p.Value);



            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;
            lbl_total_cups.Text = string.Format("Total Registros: {0:#,##0}", lista.Count());
        }

        private void Filters()
        {
            dgv.AutoGenerateColumns = false;

            if (txt_nif.Text != string.Empty)
                lista = lista.Where(z => z.nif.Contains(txt_nif.Text.ToUpper())).ToList();

            dgv.DataSource = lista;
            lbl_total_cups.Text = string.Format("Total Registros: {0:#,##0}", lista.Count());

        }

        private void Edit()
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            int c = 0;

            FrmMes13ListaNegraCUPS_Edit f = new FrmMes13ListaNegraCUPS_Edit();
            f.Text = "Editar registro";

            f.lista_negra = lista_negra;
            row = dgv.Rows[fila];

            if (row.Cells[c].Value != null)
                f.txt_nif.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_cliente.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_cups20.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_fecha_inicio.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_fecha_fin.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_comentario.Text = row.Cells[c].Value.ToString(); c++;

            f.ShowDialog();

            LoadData();
            Filters();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmMes13ListaNegraCUPS_Edit f = new FrmMes13ListaNegraCUPS_Edit();
            f.Text = "Nuevo registro";
            f.lista_negra = lista_negra;
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
                lista_negra.nif = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                lista_negra.cliente = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                lista_negra.cups20 = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                lista_negra.fecha_inicio = Convert.ToDateTime(row.Cells[c].Value.ToString()); c++;


            lista_negra.Del();

            LoadData();
            Filters();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            int c = 0;

            FrmMes13ListaNegraCUPS_Edit f = new FrmMes13ListaNegraCUPS_Edit();
            f.Text = "Editar registro";

            f.lista_negra = lista_negra;
            row = dgv.Rows[fila];

            if (row.Cells[c].Value != null)
                f.txt_nif.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_cliente.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_cups20.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_fecha_inicio.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_fecha_fin.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_comentario.Text = row.Cells[c].Value.ToString(); c++;

            f.ShowDialog();

            LoadData();
            Filters();
        }

       
    }
}
