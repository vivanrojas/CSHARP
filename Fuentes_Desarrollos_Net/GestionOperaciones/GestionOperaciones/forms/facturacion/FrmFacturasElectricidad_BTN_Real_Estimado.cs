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
    public partial class FrmFacturasElectricidad_BTN_Real_Estimado : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmFacturasElectricidad_BTN_Real_Estimado()
        {

            usage.Start("Facturación", "FrmFacturasElectricidad_BTN_Real_Estimado" ,"N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_excel_Click(object sender, EventArgs e)
        {
            EndesaBusiness.facturacion.Facturacion ff = new EndesaBusiness.facturacion.Facturacion();
            ff.InformeExcel_FacturasElectricidad_BTN_Real_Estimada(
                Convert.ToDateTime(txt_fecha_factura_desde.Value),
                Convert.ToDateTime(txt_fecha_factura_hasta.Value));
        }

        private void FrmFacturasElectricidad_BTN_Real_Estimado_Load(object sender, EventArgs e)
        {
            DateTime mesAnterior = DateTime.Now.AddMonths(-1);
            int anio = mesAnterior.Year;
            int mes = mesAnterior.Month;
            int dias_del_mes = DateTime.DaysInMonth(anio, mes);

            DateTime fh = new DateTime(anio, mes, dias_del_mes);
            DateTime fd = new DateTime(fh.Year, fh.Month, 1);

            txt_fecha_factura_desde.Value = fd;
            txt_fecha_factura_hasta.Value = fh;
        }

        private void FrmFacturasElectricidad_BTN_Real_Estimado_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmFacturasElectricidad_BTN_Real_Estimado" ,"N/A");
        }
    }
}
