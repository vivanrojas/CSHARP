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
    public partial class FrmInformeEERcs : Form
    {
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmInformeEERcs()
        {
            usage.Start("Facturación", "FrmInformeEERcs" ,"N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            EndesaBusiness.eer.InformeEERFacturas informe = new EndesaBusiness.eer.InformeEERFacturas();
            informe.InformeExcel(txt_fecha_consumo_desde.Value, txt_fecha_consumo_hasta.Value);
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void FrmInformeEERcs_Load(object sender, EventArgs e)
        {
            DateTime mesAnterior = DateTime.Now.AddMonths(-1);            
            int anio = mesAnterior.Year;
            int mes = mesAnterior.Month;
            int dias_del_mes = DateTime.DaysInMonth(anio, mes);

            DateTime fh = new DateTime(anio, mes, dias_del_mes);
            DateTime fd = new DateTime(fh.Year, fh.Month, 1);

            txt_fecha_consumo_desde.Value = fd;
            txt_fecha_consumo_hasta.Value = fh;
        }

        private void FrmInformeEERcs_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmInformeEERcs" ,"N/A");
        }
    }
}
