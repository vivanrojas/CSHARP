using EndesaBusiness.servidores;
using OfficeOpenXml;
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

namespace GestionOperaciones.forms.medida
{
    public partial class FrmAdif : Form
    {
        public FrmAdif()
        {
            int anio;
            int mes;
            int dias_del_mes;
            DateTime fh = new DateTime();
            DateTime fd = new DateTime();
            DateTime mesAnterior = new DateTime();

            InitializeComponent();

            mesAnterior = DateTime.Now.AddMonths(-1);
            anio = mesAnterior.Year;
            mes = mesAnterior.Month;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);
                                 

            fh = new DateTime(anio, mes, dias_del_mes);
            fd = new DateTime(fh.Year, fh.Month, 1);

            txtFD.Value = fd;
            txtFH.Value = fh;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FrmAdifInventario f = new FrmAdifInventario();
            f.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            FrmAdif_CS f = new FrmAdif_CS();
            f.Show();
        }

        private void btnImportar_Click(object sender, EventArgs e)
        {
            FrmAdif_Importar f = new FrmAdif_Importar();
            f.Show();
        }

        private void btnMail_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            string rutaSalida = @"c:\temp\inventarioADIF_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
            GetInventarioExcel(rutaSalida, txtFD.Value, txtFH.Value);
            Cursor.Current = Cursors.Default;
        }

        private void FrmAdif_Load(object sender, EventArgs e)
        {
            EndesaBusiness.adif.InventarioFunciones invf = new EndesaBusiness.adif.InventarioFunciones();
            invf.CompruebaInventario();
        }

        private void GetInventarioExcel(string rutaSalida, DateTime fd, DateTime fh)
        {
            string strSql;            
            int f = 0; // fila
            int c = 0; // columna
            StringBuilder cuerpo = new StringBuilder();

            EndesaBusiness.adif.AdifLotes lotes = new EndesaBusiness.adif.AdifLotes(fd, fh);

            EndesaBusiness.adif.Param p = new EndesaBusiness.adif.Param();
            EndesaBusiness.office.MailCompose mail = new EndesaBusiness.office.MailCompose();
            try
            {
                FileInfo file = new FileInfo(rutaSalida);
                ExcelPackage excelPackage = new ExcelPackage(file);
                var workSheet = excelPackage.Workbook.Worksheets.Add("ADIF");

               

                f = 1;
                c = 1;
                workSheet.Cells[f, c].Value = "CCOUNIPS"; f++;

                for(int i = 0; i < lotes.lista_CUPS13.Count; i++)
                {
                    workSheet.Cells[f, c].Value = lotes.lista_CUPS13[i]; f++;
                }                
                    
                

                
                excelPackage.Save();

                cuerpo.Append(System.Environment.NewLine);
                cuerpo.Append(DateTime.Now.Hour > 14 ? "Buenas tardes:" : "Buenos días:");
                cuerpo.Append(System.Environment.NewLine);
                cuerpo.Append("Se solicita la extracción de datos para el periodo ").Append(txtFD.Value.ToString("yyyyMM")).Append(".");
                cuerpo.Append(System.Environment.NewLine);
                cuerpo.Append("Un saludo.");

                mail.para.Add(p.GetParam("MailCurvasPara"));
                mail.cc.Add(p.GetParam("MailCurvasCC"));
                mail.asunto = "Petición de extracción de curvas y resumen para ADIF";
                mail.htmlCuerpo = cuerpo.ToString();
                mail.adjuntos.Add(file.FullName);
                mail.Show();
                file.Delete();
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
                       "GetInventarioExcel",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Error);
            }
            
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void medidaHorariaADIFAgrupadaVerticalToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void medidaADIFFueraDeInventarioToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void medidaHorariaADIFExportaciónExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmAdif_Curvas_Adif f = new FrmAdif_Curvas_Adif();
            f.Show();
        }
    }
}
