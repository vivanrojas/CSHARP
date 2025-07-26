using EndesaBusiness.factoring;
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
    public partial class FrmMes13Conversion : Form
    {

        DateTime fecha_factura_desde;
        DateTime fecha_factura_hasta;
        DateTime fecha_consumo_desde;
        DateTime fecha_consumo_hasta;
        DateTime fecha_factura_hasta_agrupadas;

        Mes12 mes12;
        public FrmMes13Conversion()
        {
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_genera_excel_Click(object sender, EventArgs e)
        {
            int num_seguimiento = Convert.ToInt32("1");
            SaveFileDialog save;


            save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.FileName = @"C:\Temp\Seguimiento_mes13_"                
                + txt_factoring.Text + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {

                Cursor.Current = Cursors.WaitCursor;
                FileInfo fileInfo = new FileInfo(save.FileName);
                if (fileInfo.Exists)
                    fileInfo.Delete();


                if (chk_solo_excel.Checked)
                    SoloGenerarExcel(save.FileName);
                else
                {
                    EndesaBusiness.factoring.SeguimientoMes13 seg =
                    new EndesaBusiness.factoring.SeguimientoMes13(num_seguimiento, Convert.ToDateTime(txtFFACTURADES.Text),
                    Convert.ToDateTime(txtFFACTURAHAS.Text),
                    Convert.ToDateTime(txt_fd_consumo.Text),
                    Convert.ToDateTime(txt_fh_consumo.Text));

                    if (num_seguimiento >= 1)
                        seg.SeguimientoMes13Agrupadas(num_seguimiento, 
                            Convert.ToDateTime(txt_fd_agrupadas.Text), Convert.ToDateTime(txt_fh_agrupadas.Text));

                    seg.SeguimientoExcel(save.FileName, false, mes12);

                    Cursor.Current = Cursors.Default;

                    MessageBox.Show("Informe de seguimiento finalizado.",
                   "Seguimiento MES13",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Information);

                }

            }
                
        }

        private void SoloGenerarExcel(string fileName)
        {

            

            EndesaBusiness.factoring.SeguimientoMes13 seg = new EndesaBusiness.factoring.SeguimientoMes13();
            seg.SeguimientoExcel(fileName, false, mes12);
        }

        private void importarAdjudicaciónToolStripMenuItem_Click(object sender, EventArgs e)
        {

            bool firstOnly = true;
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Ficheros Excel|*.xlsx";
            d.Title = "Seleccione 2 archivos de adjudicación";
            d.Multiselect = true;

            SeguimientoMes13 seg = new SeguimientoMes13();


            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (firstOnly)
                {
                    seg.BorradoTablasImportacionSeguimiento();
                    firstOnly = false;
                }

                foreach (string fileName in d.FileNames)
                {
                    seg.ImportarAdjudicacion(txt_factoring.Text, fileName);

                }
                MessageBox.Show("Importación de Adjudicaciones finalizada",
                 "Importación Finalizada",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Information);
                
                txt_factoring.Text = seg.UltimoFactoringCargado();

                lbl_total_agrupadas.Text = string.Format("Total adjudicaciones agrupadas: {0:#,##0}", seg.TotalRegistrosSeg_Agrupadas());
                lbl_total_individuales.Text = string.Format("Total adjudicaciones individuales: {0:#,##0}", seg.TotalRegistrosSeg_Individuales());
            }
        }

        private void FrmMes13Conversion_Load(object sender, EventArgs e)
        {
            SeguimientoMes13 seg = new SeguimientoMes13();
            mes12 = new Mes12();

            fecha_factura_desde = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            fecha_factura_desde = fecha_factura_desde.AddMonths(-1);

            fecha_factura_hasta = DateTime.Now;

            fecha_factura_hasta_agrupadas =
               new DateTime(Convert.ToDateTime(txtFFACTURADES.Text).Year,
               Convert.ToDateTime(txtFFACTURADES.Text).Month,
               DateTime.DaysInMonth(Convert.ToDateTime(txtFFACTURADES.Text).Year, Convert.ToDateTime(txtFFACTURADES.Text).Month));

            

            fecha_consumo_desde = fecha_factura_desde.AddMonths(-1);
            fecha_consumo_hasta = fecha_factura_desde.AddDays(-1);

            txtFFACTURADES.Text = fecha_factura_desde.ToShortDateString();
            txtFFACTURAHAS.Text = fecha_factura_hasta.ToShortDateString();

            txt_fd_consumo.Text = fecha_consumo_desde.ToShortDateString();
            txt_fh_consumo.Text = fecha_consumo_hasta.ToShortDateString();

            txt_fd_agrupadas.Text = fecha_factura_desde.ToShortDateString();
            txt_fh_agrupadas.Text = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddDays(-1).ToShortDateString();

            txt_factoring.Text = seg.UltimoFactoringCargado();



            lbl_total_facturas_mes12.Text = string.Format("Total Facturas Mes12: {0:#,##0}", mes12.dic_facturas.Count());
            lbl_total_agrupadas.Text = string.Format("Total adjudicaciones agrupadas: {0:#,##0}", seg.TotalRegistrosSeg_Agrupadas());
            lbl_total_individuales.Text = string.Format("Total adjudicaciones individuales: {0:#,##0}", seg.TotalRegistrosSeg_Individuales());

        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización Mes13";
            p.tabla = "13_param";
            p.esquemaString = "FAC";
            p.Show();
        }
    }
}
