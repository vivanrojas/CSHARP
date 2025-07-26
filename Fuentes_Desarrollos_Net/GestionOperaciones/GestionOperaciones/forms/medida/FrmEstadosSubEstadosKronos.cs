using System;
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
    public partial class FrmEstadosSubEstadosKronos : Form
    {
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        EndesaBusiness.medida.EstadosSubestadosKronos estadosKronos;
        List<EndesaEntity.medida.Pendiente> lista;
        public FrmEstadosSubEstadosKronos()
        {
            usage.Start("Medida", "FrmEstadosSubEstadosKronos", "EstadosKronos");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            
            Filters();
        }

        private void FrmEstadosSubEstadosKronos_Load(object sender, EventArgs e)
        {
            LoadData();
            Filters();
        }


        private void LoadData()
        {
            estadosKronos = new EndesaBusiness.medida.EstadosSubestadosKronos();
            lista = new List<EndesaEntity.medida.Pendiente>();

            foreach (KeyValuePair<string, EndesaEntity.medida.Pendiente> p in estadosKronos.dic)
                lista.Add(p.Value);

            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;
            lbl_total_cups.Text = string.Format("Total Registros: {0:#,##0}", lista.Count());
        }

        private void Filters()
        {
            dgv.AutoGenerateColumns = false;

            if (txt_cod_estado.Text != string.Empty)
                lista = lista.Where(z => z.cod_estado.Contains(txt_cod_estado.Text.ToUpper())).ToList();

            dgv.DataSource = lista;
            lbl_total_cups.Text = string.Format("Total Registros: {0:#,##0}", lista.Count());

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {

        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            int c = 0;

            row = dgv.Rows[fila];

            if (row.Cells[c].Value != null)
                estadosKronos.cod_estado = row.Cells[c].Value.ToString(); c++;           


            estadosKronos.Del();

            LoadData();
            Filters();
        }
    }
}
