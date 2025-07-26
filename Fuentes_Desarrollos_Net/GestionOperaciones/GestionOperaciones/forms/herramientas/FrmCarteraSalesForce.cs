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
    public partial class FrmCarteraSalesForce : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmCarteraSalesForce()
        {
            usage.Start("Herramientas", "FrmCarteraSalesForce" ,"N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void importarCarteraSalesForceCSVToolStripMenuItem_Click(object sender, EventArgs e)
        {

            EndesaBusiness.cartera.Cartera_SalesForce cartera = new EndesaBusiness.cartera.Cartera_SalesForce();

            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "csv files|*.csv";
            d.Multiselect = false;            
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)                
            {
                Cursor.Current = Cursors.WaitCursor;
                foreach (string fileName in d.FileNames)
                    cartera.ImportacionSalesForce(fileName, "ES");
                cartera.Pasa_datos_tablas_definitivas("ES");
                Cursor.Current = Cursors.Default;

               
            }
        }

        private void FrmCarteraSalesForce_Load(object sender, EventArgs e)
        {

        }

        private void importarCarteraSalesForcePortugalCSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.cartera.Cartera_SalesForce cartera = new EndesaBusiness.cartera.Cartera_SalesForce();

            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "csv files|*.csv";
            d.Multiselect = false;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                foreach (string fileName in d.FileNames)
                    cartera.ImportacionSalesForce(fileName, "PT");
                cartera.Pasa_datos_tablas_definitivas("PT");
                Cursor.Current = Cursors.Default;


            }
        }

        private void FrmCarteraSalesForce_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Herramientas", "FrmCarteraSalesForce" ,"N/A");
        }
    }
}
