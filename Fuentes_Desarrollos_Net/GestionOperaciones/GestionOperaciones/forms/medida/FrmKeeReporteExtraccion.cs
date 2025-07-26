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
    public partial class FrmKeeReporteExtraccion : Form
    {

        EndesaBusiness.medida.Kee_Extraccion_Formulas kef;
        List<EndesaEntity.medida.Kee_Extraccion_Formulas> lista = new List<EndesaEntity.medida.Kee_Extraccion_Formulas>();
        int newSortColumn;
        ListSortDirection newColumnDirection = ListSortDirection.Ascending;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmKeeReporteExtraccion()
        {
            usage.Start("Medida", "FrmKeeReporteExtraccion", "N/A");
            InitializeComponent();
        }

        private void FrmKeeReporteExtraccion_Load(object sender, EventArgs e)
        {
            kef = new EndesaBusiness.medida.Kee_Extraccion_Formulas();
            InitListBox();            
            lista = kef.list;
            ActualizaListBox_Sufijo_CUPS();
            Carga_DGV();            
        }

        private void Carga_DGV()
        {
            lbl_registros.Text = string.Format("Registros: {0:#,##0}", lista.Count());
            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;

            
        }

       

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void InitListBox()
        {

            for (int i = listBoxExtraccion.Items.Count - 1; i == 0; i--)
            {
                listBoxExtraccion.Items.RemoveAt(i);
            }
            
            foreach (string p in kef.lista_extraccion)
                listBoxExtraccion.Items.Add(p);

            //for (int i = listBoxSufijo_CUPS22.Items.Count - 1; i == 0; i--)
            //{
            //    listBoxSufijo_CUPS22.Items.RemoveAt(i);
            //}

            //foreach (string p in kef.lista_sufijos_cups22)
            //    listBoxSufijo_CUPS22.Items.Add(p);

        }

        private void listBoxExtraccion_SelectedIndexChanged(object sender, EventArgs e)
        {
            Actualiza_Lista();
            ActualizaListBox_Sufijo_CUPS();


        }

        private void btn_po1011_Click(object sender, EventArgs e)
        {
            EndesaBusiness.medida.Kee_Reporte_Extraccion_CH kee = new EndesaBusiness.medida.Kee_Reporte_Extraccion_CH(true);
            kee.InformePO1011(lista);
        }

        private void listBoxSufijo_CUPS22_SelectedIndexChanged(object sender, EventArgs e)
        {
            Actualiza_Lista();            
        }


        private void Actualiza_Lista()
        {
            List<EndesaEntity.medida.Kee_Extraccion_Formulas> lista_tmp = new List<EndesaEntity.medida.Kee_Extraccion_Formulas>();


            lista = kef.list;

            if (listBoxExtraccion.SelectedItems.Count > 0)
            {

                foreach (Object value in listBoxExtraccion.SelectedItems)
                {
                    List<EndesaEntity.medida.Kee_Extraccion_Formulas> l = lista.Where(z => z.extraccion == value.ToString()).Select(z => z).ToList();
                    foreach (EndesaEntity.medida.Kee_Extraccion_Formulas p in l)
                        lista_tmp.Add(p);

                }

                
                lista = lista_tmp;

            }

            if (listBoxSufijo_CUPS22.SelectedItems.Count > 0)
            {

                List<EndesaEntity.medida.Kee_Extraccion_Formulas> l = new List<EndesaEntity.medida.Kee_Extraccion_Formulas>();

                foreach (Object value in listBoxSufijo_CUPS22.SelectedItems)
                {

                    foreach (EndesaEntity.medida.Kee_Extraccion_Formulas p in lista)                        
                    {
                        if(p.cups22 != null)
                            if (p.cups22.Substring(20,2) == value.ToString())
                                if (!l.Exists(z => z.cups20 == p.cups20))
                                    l.Add(p);
                    }

                }

                lista = l;

            }

            Carga_DGV();
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

        private List<EndesaEntity.medida.Kee_Extraccion_Formulas> 
            OrdenaColumna(string columna, ListSortDirection direccion)
        {
            List<EndesaEntity.medida.Kee_Extraccion_Formulas> l =
                new List<EndesaEntity.medida.Kee_Extraccion_Formulas>();

            switch (columna)
            {
                case "cups20":
                    if(direccion == ListSortDirection.Ascending)
                        l = lista.OrderBy(z => z.cups20).ToList();
                    else
                        l = lista.OrderByDescending(z => z.cups20).ToList();
                    break;

                case "cups22":
                    if (direccion == ListSortDirection.Ascending)
                        l = lista.OrderBy(z => z.cups22).ToList();
                    else
                        l = lista.OrderByDescending(z => z.cups22).ToList();
                    break;
                case "fecha_sol_desde":
                    if (direccion == ListSortDirection.Ascending)
                        l = lista.OrderBy(z => z.fecha_sol_desde).ToList();
                    else
                        l = lista.OrderByDescending(z => z.fecha_sol_desde).ToList();
                    break;
                case "fecha_sol_hasta":
                    if (direccion == ListSortDirection.Ascending)
                        l = lista.OrderBy(z => z.fecha_sol_hasta).ToList();
                    else
                        l = lista.OrderByDescending(z => z.fecha_sol_hasta).ToList();
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

                case "tipo":
                    if (direccion == ListSortDirection.Ascending)
                        l = lista.OrderBy(z => z.tipo).ToList();
                    else
                        l = lista.OrderByDescending(z => z.tipo).ToList();
                    break;

                case "fuente":
                    if (direccion == ListSortDirection.Ascending)
                        l = lista.OrderBy(z => z.fuente).ToList();
                    else
                        l = lista.OrderByDescending(z => z.fuente).ToList();
                    break;

                case "extraccion":
                    if (direccion == ListSortDirection.Ascending)
                        l = lista.OrderBy(z => z.extraccion).ToList();
                    else
                        l = lista.OrderByDescending(z => z.extraccion).ToList();
                    break;

                case "fecha_mod_archivo":
                    if (direccion == ListSortDirection.Ascending)
                        l = lista.OrderBy(z => z.fecha_mod_archivo).ToList();
                    else
                        l = lista.OrderByDescending(z => z.fecha_mod_archivo).ToList();
                    break;

                    
            }

            return l;
                
        }

        private void cmdExcel_Click_1(object sender, EventArgs e)
        {
            EndesaBusiness.medida.Kee_Reporte_Extraccion_CH kee = new EndesaBusiness.medida.Kee_Reporte_Extraccion_CH(true);
            kee.InformeExcel(lista);
        }

        private void ayudaToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void ActualizaListBox_Sufijo_CUPS()
        {
            listBoxSufijo_CUPS22.Items.Clear();

            List<string> o = new List<string>();
            foreach (EndesaEntity.medida.Kee_Extraccion_Formulas p in lista)
            {
                if(p.cups22 != null)
                    if (!o.Exists(z => z == p.cups22.Substring(20, 2)))
                        o.Add(p.cups22.Substring(20, 2));
            }                
                


            foreach (string p in o)
                listBoxSufijo_CUPS22.Items.Add(p);
        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "kee_param";
            p.esquemaString = "MED";
            p.cabecera = "Parámetros KEE (Kronos Endesa Energía)";
            p.Show();
        }

        private void FrmKeeReporteExtraccion_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Medida", "FrmKeeReporteExtraccion", "N/A");
        }
    }
}
