using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.contratacion.gas
{
    public partial class FrmFechasInforme : Form
    {
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public bool cancelado { get; set; }
        public FrmFechasInforme()
        {
            cancelado = false;
            InitializeComponent();
        }

        private void FrmFechasInforme_Load(object sender, EventArgs e)
        {
            DateTime fd;
            DateTime fh;
            DateTime mesAnterior;
            int anio;
            int mes;
            int dias_del_mes;

            // InitializeComponent();

            usage.Start("Contratación", "FrmFechasInforme" ,"N/A");

            mesAnterior = DateTime.Now.AddMonths(-1);
            anio = mesAnterior.Year;
            mes = mesAnterior.Month;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            fh = new DateTime(anio, mes, dias_del_mes);
            fd = new DateTime(fh.Year, fh.Month, 1);

            txt_fd.Value = fd;
            txt_fh.Value = fh;

            
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            cancelado = true;
            this.Close();
        }

        private void FrmFechasInforme_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Contratación", "FrmFechasInforme" ,"N/A");
        }
    }
}
