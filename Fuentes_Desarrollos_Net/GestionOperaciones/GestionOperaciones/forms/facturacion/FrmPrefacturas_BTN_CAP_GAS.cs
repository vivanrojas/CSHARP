using EndesaBusiness.contratacion.eexxi;
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
    public partial class FrmPrefacturas_BTN_CAP_GAS : Form
    {
        EndesaBusiness.facturacion.puntos_calculo_btn.Calculo_Prefacturas_BTN btn;
        EndesaBusiness.facturacion.puntos_calculo_btn.Ajustes ajustes;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        int newSortColumn;
        ListSortDirection newColumnDirection = ListSortDirection.Ascending;
        public FrmPrefacturas_BTN_CAP_GAS()
        {
            usage.Start("Facturación", "FrmPrefacturas_BTN_CAP_GAS" ,"N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void importarExcelExtracciónDeAjustesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = false;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    ajustes.CargaExcel(fileName);
                    if (!ajustes.hayError)
                    {
                        MessageBox.Show("Importación completada correctamente",
                            "Importar Excel",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);

                    }
                    else
                    {
                        MessageBox.Show(ajustes.descripcion_error,
                           "Importar Excel",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Information);
                    }

                }
            }
        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización Prefacturas BTN - CAP GAS";
            p.tabla = "lpc_btn_param";
            p.esquemaString = "FAC";
            p.ShowDialog();
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void FrmPrefacturas_BTN_CAP_GAS_Load(object sender, EventArgs e)
        {
            btn = new EndesaBusiness.facturacion.puntos_calculo_btn.Calculo_Prefacturas_BTN();
            ajustes = new EndesaBusiness.facturacion.puntos_calculo_btn.Ajustes();

            Cursor.Current = Cursors.WaitCursor;
            LoadData();
            Cursor.Current = Cursors.Default;

        }

        private void LoadData()
        {
            btn.Proceso();
            dgv.AutoGenerateColumns = false;
            dgv.DataSource = btn.dic.Values.ToList().OrderBy(z => z.cpe).ToList();
            lbl_total_cups.Text = string.Format("Total Registros: {0:#,##0}", dgv.RowCount);
        }

        private void cmdExcel_Click(object sender, EventArgs e)
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
                btn.RellenaExcel(save.FileName);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }
        }


        private List<EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo>
    OrdenaColumna(string columna, ListSortDirection direccion)
        {
            List<EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> l =
                new List<EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo>();

            switch (columna)
            {
                case "cpe":
                    if (direccion == ListSortDirection.Ascending)
                        l = btn.dic.Values.ToList().OrderBy(z => z.cpe).ToList();
                    else
                        l = btn.dic.Values.ToList().OrderByDescending(z => z.cpe).ToList();
                    break;

                case "f_desde":
                    if (direccion == ListSortDirection.Ascending)
                        l = btn.dic.Values.ToList().OrderBy(z => z.f_desde).ToList();
                    else
                        l = btn.dic.Values.ToList().OrderByDescending(z => z.f_desde).ToList();
                    break;

                case "f_hasta":
                    if (direccion == ListSortDirection.Ascending)
                        l = btn.dic.Values.ToList().OrderBy(z => z.f_hasta).ToList();
                    else
                        l = btn.dic.Values.ToList().OrderByDescending(z => z.f_hasta).ToList();
                    break;
                case "consumo":
                    if (direccion == ListSortDirection.Ascending)
                        l = btn.dic.Values.ToList().OrderBy(z => z.consumo).ToList();
                    else
                        l = btn.dic.Values.ToList().OrderByDescending(z => z.consumo).ToList();
                    break;

                case "perfil":
                    if (direccion == ListSortDirection.Ascending)
                        l = btn.dic.Values.ToList().OrderBy(z => z.perfil).ToList();
                    else
                        l = btn.dic.Values.ToList().OrderByDescending(z => z.perfil).ToList();
                    break;

                case "calendario":
                    if (direccion == ListSortDirection.Ascending)
                        l = btn.dic.Values.ToList().OrderBy(z => z.calendario).ToList();
                    else
                        l = btn.dic.Values.ToList().OrderByDescending(z => z.calendario).ToList();
                    break;
                case "tarifa":
                    if (direccion == ListSortDirection.Ascending)
                        l = btn.dic.Values.ToList().OrderBy(z => z.tarifa).ToList();
                    else
                        l = btn.dic.Values.ToList().OrderByDescending(z => z.tarifa).ToList();
                    break;
                case "pct_aplicacion":
                    if (direccion == ListSortDirection.Ascending)
                        l = btn.dic.Values.ToList().OrderBy(z => z.pct_aplicacion).ToList();
                    else
                        l = btn.dic.Values.ToList().OrderByDescending(z => z.pct_aplicacion).ToList();
                    break;                


            }

            return l;

        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

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

        private void rellenarPlantillaDeCálculoToolStripMenuItem_Click(object sender, EventArgs e)
        {

            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsm";
            save.Filter = "Ficheros xslm (*.xlsm)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                btn.RellenaExcel(save.FileName);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }

            
        }

        private void FrmPrefacturas_BTN_CAP_GAS_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmPrefacturas_BTN_CAP_GAS" ,"N/A");
        }
    }
}
