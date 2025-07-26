using EndesaBusiness.factoring;
using GestionOperaciones.forms.contratacion.gas;
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

namespace GestionOperaciones.forms.medida
{
    public partial class FrmAdifCierresEnergia : Form
    {


        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        EndesaBusiness.adif.CierresEnergia cierres;
        List<EndesaEntity.medida.CierresEnergia> lista;

        int newSortColumn;
        ListSortDirection newColumnDirection = ListSortDirection.Ascending;

        public FrmAdifCierresEnergia()
        {
            usage.Start("Medida", "FrmAdifCierresEnergia", "Cierres de energía");
            InitializeComponent();

        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmAdifCierresEnergia_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void FrmAdifCierresEnergia_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Medida", "FrmAdifCierresEnergia", "Cierres de energía");
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            LoadData();
            Filters();
        }

        private void LoadData()
        {
            cierres = new EndesaBusiness.adif.CierresEnergia();

            lista = new List<EndesaEntity.medida.CierresEnergia>();

            foreach (KeyValuePair<string, List<EndesaEntity.medida.CierresEnergia>> p in cierres.dic)
            {
                for (int i = 0; i < p.Value.Count(); i++)
                    lista.Add(p.Value[i]);
            }


            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;
            lbl_total_cups.Text = string.Format("Total Registros: {0:#,##0}", lista.Count());
        }

        private void Filters()
        {
            dgv.AutoGenerateColumns = false;

            if(txt_cups20.Text != string.Empty)            
                lista = lista.Where(z => z.cups20.Contains(txt_cups20.Text.ToUpper())).ToList();

            dgv.DataSource = lista;
            lbl_total_cups.Text = string.Format("Total Registros: {0:#,##0}", lista.Count()); 

        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            Edit();
        }

        private void dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Edit();
        }

        private void Edit()
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            int c = 0;

            FrmAdifCierresEnergia_Edit f = new FrmAdifCierresEnergia_Edit();
            f.Text = "Editar cierre";
            f.cierres = cierres;
            f.txt_cups20.Enabled = false;

            row = dgv.Rows[fila];

            if (row.Cells[c].Value != null)
                f.txt_cups20.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_fd.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_fh.Text = row.Cells[c].Value.ToString(); c++;

            f.ShowDialog();
            LoadData();
            Filters();

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            forms.medida.FrmAdifCierresEnergia_Edit f = new FrmAdifCierresEnergia_Edit();
            f.Text = "Nuevo cierre";
            f.cierres = cierres;
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
                cierres.cups20 = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                cierres.fecha_desde = Convert.ToDateTime(row.Cells[c].Value); c++;

            

            cierres.Del();

            LoadData();
            Filters();
        }

        private List<EndesaEntity.medida.CierresEnergia> OrdenaColumna(string columna, ListSortDirection direccion)
        {
            List<EndesaEntity.medida.CierresEnergia> l =
                new List<EndesaEntity.medida.CierresEnergia>();

            switch (columna)
            {
                case "cups20":
                    if (direccion == ListSortDirection.Ascending)
                        l = lista.OrderBy(z => z.cups20).ToList();  
                    else
                        l = lista.OrderByDescending(z => z.cups20).ToList();
                    break;

                case "fecha_desde":
                    if (direccion == ListSortDirection.Ascending)
                        l = lista.OrderBy(z => z.fecha_desde).ToList();
                    else
                        l = lista.OrderByDescending(z => z.fecha_desde).ToList();
                    break;

                case "fecha_hasta":
                    if (direccion == ListSortDirection.Ascending)
                        l = lista.OrderBy(z => z.fecha_hasta).ToList();
                    else
                        l = lista.OrderByDescending(z => z.fecha_hasta).ToList();
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
        
    }
}
