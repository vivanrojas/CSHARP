using EndesaBusiness.logs;
using EndesaBusiness.servidores;
using EndesaBusiness.sharePoint;
using EndesaBusiness.utilidades;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmMes13Prevision : Form
    {
        private EndesaBusiness.utilidades.Param param;
        EndesaBusiness.logs.Log ficheroLog;
        public FrmMes13Prevision()
        {
            InitializeComponent();
            ficheroLog = new EndesaBusiness.logs.Log(Environment.CurrentDirectory, "logs", "FAC_Mes13_Prevision");
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmMes13Prevision_Load(object sender, EventArgs e)
        {
            EndesaBusiness.factoring.FechasExtracciones extracciones =
                new EndesaBusiness.factoring.FechasExtracciones();

            param = new Param("ff_param", MySQLDB.Esquemas.FAC);

            txt_factoring.Text = DateTime.Now.ToString("yyyyMM");
            ActualizaBloques();

            txt_importe_facturas_individuales.Text = String.Format(System.Globalization.CultureInfo.CurrentCulture, "{0:C2}", Convert.ToDouble(param.GetValue("importe_facturas_individuales")));
            txt_importe_facturas_agrupadas.Text = String.Format(System.Globalization.CultureInfo.CurrentCulture, "{0:C2}", Convert.ToDouble(param.GetValue("importe_facturas_agrupadas")));

            this.lbl_obligaciones_cobradas.Text = string.Format("Oblig. Cobradas actualización: {0}", 
                extracciones.UltimaActualizacionObligCobradas());
            this.lbl_obligaciones.Text = string.Format("Oblig. Deuda actualización: {0}",
                extracciones.UltimaActualizacionObligDeuda());

        }

        private void ActualizaBloques()
        {
            int anio = Convert.ToInt32(txt_factoring.Text.Substring(0, 4));
            int mes = Convert.ToInt32(txt_factoring.Text.Substring(4, 2));

            DateTime fd = new DateTime(anio, mes, 1);
            DateTime fh = new DateTime(fd.Year, fd.Month, DateTime.DaysInMonth(fd.Year, fd.Month));

            txt_fd_0.Value = fd.AddMonths(-12);
            txt_fh_0.Value = fh.AddMonths(-12);
            txt_fd_1.Value = fd.AddMonths(-4);
            txt_fh_1.Value = fd.AddMonths(-3).AddDays(-1);
            txt_fd_2.Value = fd.AddMonths(-3);
            txt_fh_2.Value = fd.AddMonths(-2).AddDays(-1);
            txt_fd_3.Value = fd.AddMonths(-2);
            txt_fh_3.Value = fd.AddMonths(-1).AddDays(-1);
                       

        }

        
        private void txt_factoring_Leave(object sender, EventArgs e)
        {
            ActualizaBloques();
        }

        private void btnEdit_d_Click(object sender, EventArgs e)
        {
            txt_fd_0.Enabled = true; 
            txt_fh_0.Enabled = true;
            txt_fd_1.Enabled = true;
            txt_fh_1.Enabled = true;
            txt_fd_2.Enabled = true;
            txt_fh_2.Enabled = true;
            txt_fd_3.Enabled = true;
            txt_fh_3.Enabled = true;
            txt_importe_facturas_individuales.Enabled = true;
            txt_importe_facturas_agrupadas.Enabled = true; 
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void btn_generar_prevision_Click(object sender, EventArgs e)
        {

            double importe_min_individuales = Convert.ToDouble(txt_importe_facturas_individuales.Text.Replace("€", ""));
            double importe_min_agrupadas = Convert.ToDouble(txt_importe_facturas_agrupadas.Text.Replace("€", ""));

            SaveFileDialog save;

            save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.FileName = @"C:\Temp\factoring_"
                + txt_factoring.Text + "_estimacion_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {

                FileInfo fileInfo = new FileInfo(save.FileName);
                if (fileInfo.Exists)
                    fileInfo.Delete();



                EndesaBusiness.factoring.Prevision prevision = new EndesaBusiness.factoring.Prevision();

                Dictionary<string, List<EndesaEntity.factoring.CalendarioFactoring>> dic_calendario =
                    new Dictionary<string, List<EndesaEntity.factoring.CalendarioFactoring>>();


                EndesaEntity.factoring.CalendarioFactoring c;
                List<EndesaEntity.factoring.CalendarioFactoring> lista = new List<EndesaEntity.factoring.CalendarioFactoring>();

                c = new EndesaEntity.factoring.CalendarioFactoring();
                c.factoring = txt_factoring.Text;                
                c.consumos_desde = txt_fd_0.Value;
                c.consumos_hasta = txt_fh_0.Value;
                c.facturas_desde = c.consumos_desde.AddMonths(1);
                c.facturas_hasta = new DateTime(c.facturas_desde.Year, c.facturas_desde.Month,
                    DateTime.DaysInMonth(c.facturas_desde.Year, c.facturas_desde.Month));
                c.bloque = 0;
                c.importe_min_factura = importe_min_individuales;
                c.importe_min_factura_agrupada = importe_min_agrupadas;
                lista.Add(c);

                c = new EndesaEntity.factoring.CalendarioFactoring();
                c.factoring = txt_factoring.Text;
                c.consumos_desde = txt_fd_1.Value;
                c.consumos_hasta = txt_fh_1.Value;
                c.facturas_desde = c.consumos_desde.AddMonths(1);
                c.facturas_hasta = new DateTime(c.facturas_desde.Year, c.facturas_desde.Month,
                    DateTime.DaysInMonth(c.facturas_desde.Year, c.facturas_desde.Month));
                c.bloque = 1;
                c.importe_min_factura = importe_min_individuales;
                c.importe_min_factura_agrupada = importe_min_agrupadas;
                lista.Add(c);

                c = new EndesaEntity.factoring.CalendarioFactoring();
                c.factoring = txt_factoring.Text;
                c.consumos_desde = txt_fd_2.Value;
                c.consumos_hasta = txt_fh_2.Value;
                c.facturas_desde = c.consumos_desde.AddMonths(1);
                c.facturas_hasta = new DateTime(c.facturas_desde.Year, c.facturas_desde.Month,
                    DateTime.DaysInMonth(c.facturas_desde.Year, c.facturas_desde.Month));
                c.bloque = 2;
                c.importe_min_factura = importe_min_individuales;
                c.importe_min_factura_agrupada = importe_min_agrupadas;
                lista.Add(c);

                c = new EndesaEntity.factoring.CalendarioFactoring();
                c.factoring = txt_factoring.Text;
                c.consumos_desde = txt_fd_3.Value;
                c.consumos_hasta = txt_fh_3.Value;
                c.facturas_desde = c.consumos_desde.AddMonths(1);
                c.facturas_hasta = new DateTime(c.facturas_desde.Year, c.facturas_desde.Month,
                    DateTime.DaysInMonth(c.facturas_desde.Year, c.facturas_desde.Month));
                c.bloque = 3;
                c.importe_min_factura = importe_min_individuales;
                c.importe_min_factura_agrupada = importe_min_agrupadas;
                lista.Add(c);


                dic_calendario.Add(txt_factoring.Text, lista);
                prevision.dic = dic_calendario;
                prevision.CreaPrevision(fileInfo.FullName);
            }

        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización Mes13";
            p.tabla = "ff_param";
            p.esquemaString = "FAC";
            p.Show();
        }

        private void listaNegraToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmMes13ListaNegra f = new FrmMes13ListaNegra();
            f.Show();
        }

        private void listaNegraCUPSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmMes13ListaNegraCUPS f = new FrmMes13ListaNegraCUPS();
            f.Show();
        }
    }
}
