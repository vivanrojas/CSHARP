using GestionOperaciones.forms.contratacion.gas;
using GestionOperaciones.forms.medida;
using OfficeOpenXml.LoadFunctions.Params;
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
    public partial class FrmMes13ListaNegra : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        EndesaBusiness.factoring.Lista_Negra lista_negra;
        List<EndesaEntity.factoring.ListaNegra> lista;

        int newSortColumn;
        ListSortDirection newColumnDirection = ListSortDirection.Ascending;
        public FrmMes13ListaNegra()
        {

            usage.Start("Facturación", "FrmMes13ListaNegra", "Lista Negra");
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

        private void FrmMes13ListaNegra_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmMes13ListaNegra", "Lista Negra");
        }

        private void FrmMes13ListaNegra_Load(object sender, EventArgs e)
        {
            LoadData();
        }


        private void LoadData()
        {
            lista_negra = new EndesaBusiness.factoring.Lista_Negra();

            lista = new List<EndesaEntity.factoring.ListaNegra>();

            foreach (KeyValuePair<string, EndesaEntity.factoring.ListaNegra> p in lista_negra.dic)
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

        private List<EndesaEntity.factoring.ListaNegra> OrdenaColumna(string columna, ListSortDirection direccion)
        {
            List<EndesaEntity.factoring.ListaNegra> l =
                new List<EndesaEntity.factoring.ListaNegra>();

            switch (columna)
            {
                case "nif":
                    if (direccion == ListSortDirection.Ascending)
                        l = lista_negra.dic.Values.OrderBy(z => z.nif).ToList();
                    else
                        l = lista_negra.dic.Values.OrderByDescending(z => z.nif).ToList();
                    break;

                case "cliente":
                    if (direccion == ListSortDirection.Ascending)
                        l = lista_negra.dic.Values.OrderBy(z => z.cliente).ToList();
                    else
                        l = lista_negra.dic.Values.OrderByDescending(z => z.cliente).ToList();
                    break;
            }

            return l;
        }

        private void dgv_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dgv.Columns[e.ColumnIndex].SortMode != DataGridViewColumnSortMode.NotSortable)
            {
                if (e.ColumnIndex == newSortColumn)
                {
                    if (newColumnDirection == ListSortDirection.Ascending)
                        newColumnDirection = ListSortDirection.Descending;
                    else
                        newColumnDirection = ListSortDirection.Ascending;
                }

                newSortColumn = e.ColumnIndex;

                dgv.AutoGenerateColumns = false;
                dgv.DataSource = OrdenaColumna(dgv.Columns[e.ColumnIndex].DataPropertyName, newColumnDirection); ;


            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            Edit();
        }


        private void Edit()
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            int c = 0;

            FrmMes13ListaNegra_Edit f = new FrmMes13ListaNegra_Edit();
            f.Text = "Editar registro";

            f.lista_negra = lista_negra;
            row = dgv.Rows[fila];

            if (row.Cells[c].Value != null)
                f.txt_nif.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_cliente.Text = row.Cells[c].Value.ToString(); c++;

            f.ShowDialog();

            LoadData();
            Filters();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmMes13ListaNegra_Edit f = new FrmMes13ListaNegra_Edit();
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


            lista_negra.Del();

            LoadData();
            Filters();
        }
    }




}
