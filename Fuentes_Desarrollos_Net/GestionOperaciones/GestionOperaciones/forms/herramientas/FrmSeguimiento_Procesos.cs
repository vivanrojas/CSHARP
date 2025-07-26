using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.herramientas
{
    public partial class FrmSeguimiento_Procesos : Form
    {
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;
        List<EndesaEntity.herramientas.Seguimiento_Procesos> lista = new List<EndesaEntity.herramientas.Seguimiento_Procesos>();
        int newSortColumn;
        ListSortDirection newColumnDirection = ListSortDirection.Ascending;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmSeguimiento_Procesos()
        {
            usage.Start("Herramientas", "FrmSeguimiento_Procesos", this.Name);
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmSeguimiento_Procesos_Load(object sender, EventArgs e)
        {
            ss_pp = new EndesaBusiness.utilidades.Seguimiento_Procesos();
            InitListBox();
            lista = ss_pp.dic.Values.ToList();
            Carga_DGV();
            LoadDataDetail();
        }

        private void InitListBox()
        {
            listBoxArea.Items.Clear();

            foreach (string p in ss_pp.lista_areas)
                listBoxArea.Items.Add(p);

            listBoxProceso.Items.Clear();

            foreach (string p in ss_pp.lista_procesos)
                listBoxProceso.Items.Add(p);

        }

        private void Carga_DGV()
        {
            //lbl_registros.Text = string.Format("Registros: {0:#,##0}", lista.Count());
            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if(row.Cells[5].Value != null)
                    if (Convert.ToString(row.Cells[5].Value) == "OK")
                        row.DefaultCellStyle.BackColor = Color.LightGreen;

                if (Convert.ToDateTime(row.Cells[4].Value).Date == DateTime.MinValue)
                    row.DefaultCellStyle.BackColor = Color.LightYellow;
            }
        }

        private void listBoxProceso_SelectedIndexChanged(object sender, EventArgs e)
        {
            Actualiza_Lista();
            
        }

        private void Actualiza_Lista()
        {
            List<string> lista_procesos = new List<string>();
            List<EndesaEntity.herramientas.Seguimiento_Procesos> lista_tmp = new List<EndesaEntity.herramientas.Seguimiento_Procesos>();

            lista = ss_pp.dic.Values.ToList();

            if (listBoxArea.SelectedItems.Count > 0)
            {
                foreach (Object value in listBoxArea.SelectedItems)
                {
                    //List<string> l = ss_pp.lista.Where(z => z.area == value.ToString()).Select(z => z.proceso).ToList();
                    List<EndesaEntity.herramientas.Seguimiento_Procesos> l =
                        lista.Where(z => z.area == value.ToString()).Select(z => z).ToList();
                    foreach (EndesaEntity.herramientas.Seguimiento_Procesos s in l)
                        lista_tmp.Add(s);
                }

                lista = lista_tmp;
            }

            if (listBoxProceso.SelectedItems.Count > 0)
            {
                List<EndesaEntity.herramientas.Seguimiento_Procesos> l = new List<EndesaEntity.herramientas.Seguimiento_Procesos>();

                foreach (Object value in listBoxProceso.SelectedItems)
                {


                    foreach (EndesaEntity.herramientas.Seguimiento_Procesos p in lista)
                    {
                        if (p.proceso != null)
                            if (p.proceso == value.ToString())
                                l.Add(p);
                    }
                }

                lista = l;

            }

            Carga_DGV();
        }

        private void listBoxArea_SelectedIndexChanged(object sender, EventArgs e)
        {
            Actualiza_Lista();
            ActualizaListBox_Procesos();
        }

        private void ActualizaListBox_Procesos()
        {
            listBoxProceso.Items.Clear();

            List<string> o = new List<string>();
            foreach (EndesaEntity.herramientas.Seguimiento_Procesos p in lista)
            {
                if (p.proceso != null)
                    if (!o.Exists(z => z == p.proceso))
                        o.Add(p.proceso);
            }

            foreach (string p in o)
                listBoxProceso.Items.Add(p);
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

                lista = OrdenaColumna(dgv.Columns[e.ColumnIndex].DataPropertyName, newColumnDirection);
                Carga_DGV();
            }
        }
            
        private List<EndesaEntity.herramientas.Seguimiento_Procesos>
           OrdenaColumna(string columna, ListSortDirection direccion)
            {
                List<EndesaEntity.herramientas.Seguimiento_Procesos> l =
                    new List<EndesaEntity.herramientas.Seguimiento_Procesos>();

                switch (columna)
                {
                    case "area":
                        if (direccion == ListSortDirection.Ascending)
                            l = lista.OrderBy(z => z.area).ToList();
                        else
                            l = lista.OrderByDescending(z => z.area).ToList();
                        break;

                    case "proceso":
                        if (direccion == ListSortDirection.Ascending)
                            l = lista.OrderBy(z => z.proceso).ToList();
                        else
                            l = lista.OrderByDescending(z => z.proceso).ToList();
                        break;
                    //case "paso":
                    //    if (direccion == ListSortDirection.Ascending)
                    //        l = lista.OrderBy(z => z.paso).ToList();
                    //    else
                    //        l = lista.OrderByDescending(z => z.paso).ToList();
                    //    break;
                    case "descripcion":
                        if (direccion == ListSortDirection.Ascending)
                            l = lista.OrderBy(z => z.descripcion).ToList();
                        else
                            l = lista.OrderByDescending(z => z.descripcion).ToList();
                        break;

                    case "fecha_inicio":
                        if (direccion == ListSortDirection.Ascending)
                            l = lista.OrderBy(z => z.fecha_inicio).ToList();
                        else
                            l = lista.OrderByDescending(z => z.fecha_inicio).ToList();
                        break;
                    case "fecha_fin":
                        if (direccion == ListSortDirection.Ascending)
                            l = lista.OrderBy(z => z.fecha_fin).ToList();
                        else
                            l = lista.OrderByDescending(z => z.fecha_fin).ToList();
                        break;

                    case "comentario":
                        if (direccion == ListSortDirection.Ascending)
                            l = lista.OrderBy(z => z.comentario).ToList();
                        else
                            l = lista.OrderByDescending(z => z.comentario).ToList();
                        break;

                    case "ejecucion":
                        if (direccion == ListSortDirection.Ascending)
                            l = lista.OrderBy(z => z.ejecucion).ToList();
                        else
                            l = lista.OrderByDescending(z => z.ejecucion).ToList();
                        break;

                    case "tarea":
                        if (direccion == ListSortDirection.Ascending)
                            l = lista.OrderBy(z => z.tarea).ToList();
                        else
                            l = lista.OrderByDescending(z => z.tarea).ToList();
                        break;

                    case "contacto":
                        if (direccion == ListSortDirection.Ascending)
                            l = lista.OrderBy(z => z.contacto).ToList();
                        else
                            l = lista.OrderByDescending(z => z.contacto).ToList();
                        break;


                }

                return l;

            }

       
        

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            ss_pp = new EndesaBusiness.utilidades.Seguimiento_Procesos();
            InitListBox();
            lista = ss_pp.dic.Values.ToList();
            Carga_DGV();
        }

        private void LoadDataDetail()
        {
            dgvd.AutoGenerateColumns = false;
            dgvd.DataSource = ss_pp.dic.FirstOrDefault(z => z.Key == ss_pp.dic.First().Key).Value.detalle.ToList();
        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización Seguimiento Procesos";
            p.tabla = "global_param";
            p.esquemaString = "AUX";
            p.Show();
        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dgv_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewCell area = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[0];

                DataGridViewCell proceso = (DataGridViewCell)
                   dgv.Rows[e.RowIndex].Cells[1];

                dgvd.AutoGenerateColumns = false;
                EndesaEntity.herramientas.Seguimiento_Procesos p;
                if (ss_pp.dic.TryGetValue(area.Value.ToString() + "_" + proceso.Value.ToString(), out p))
                    dgvd.DataSource = p.detalle.ToList();
                else
                    dgvd.DataSource = null;

                foreach (DataGridViewRow row in dgvd.Rows)
                {
                    if (row.Cells[5].Value != null)
                        if (Convert.ToString(row.Cells[5].Value) == "OK")
                            row.DefaultCellStyle.BackColor = Color.LightGreen;

                    if (Convert.ToDateTime(row.Cells[4].Value).Date == DateTime.MinValue)
                        row.DefaultCellStyle.BackColor = Color.LightYellow;
                    
                }

            }
        }

        private void FrmSeguimiento_Procesos_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Herramientas", "FrmSeguimiento_Procesos", "N/A");
        }
    }
}
