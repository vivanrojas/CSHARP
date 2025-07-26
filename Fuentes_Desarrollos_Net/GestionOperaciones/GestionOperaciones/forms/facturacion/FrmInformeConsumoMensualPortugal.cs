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
    public partial class FrmInformeConsumoMensualPortugal : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmInformeConsumoMensualPortugal()
        {

            usage.Start("Facturación", "FrmInformeConsumoMensualPortugal" ,"N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            int c = 1;
            int f = 1;
            SaveFileDialog save;
            try
            {

            }
            catch (Exception ex)
            {

            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void FrmInformeConsumoMensualPortugal_Load(object sender, EventArgs e)
        {
            DateTime mesAnterior = DateTime.Now.AddMonths(-1);
            int anio = mesAnterior.Year;
            int mes = mesAnterior.Month;
            int dias_del_mes = DateTime.DaysInMonth(anio, mes);

            DateTime fh = new DateTime(anio, mes, dias_del_mes);
            DateTime fd = new DateTime(fh.Year, fh.Month, 1);

            txt_fecha_consumo_desde.Value = new DateTime(2021,01,01);
            txt_fecha_consumo_hasta.Value = fh;
        }

        private void FrmInformeConsumoMensualPortugal_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmInformeConsumoMensualPortugal" ,"N/A");
        }
    }
}
