//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Spreadsheet;


using EndesaBusiness.medida.Redshift;
using EndesaEntity;
using EndesaEntity.global;
using EndesaEntity.medida;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace GestionOperaciones.forms.medida
{
    public partial class FrmCurvas : Form
    {

        EndesaBusiness.medida.ExcelCUPS ex;
        private List<string> lista_drivers = new List<string>();

        List<EndesaEntity.medida.PuntoSuministro> lc = new List<EndesaEntity.medida.PuntoSuministro>();
        EndesaBusiness.logs.Log ficheroLog = new
            EndesaBusiness.logs.Log(Environment.CurrentDirectory, "logs", System.Environment.UserName);

        List<EndesaEntity.medida.CurvaCuartoHorariaInformes> list_cc;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmCurvas()
        {

            usage.Start("Medida", "FrmCurvas", "N/A");
            InitializeComponent();

        }

        private void FrmCurvas_Load(object sender, EventArgs e)
        {
            List<string> lista_drivers = new List<string>();
            DateTime fd = new DateTime();
            DateTime fh = new DateTime();
            int anio;
            int mes;
            int dias_del_mes;
            radio_facturada.Checked = true;

            EndesaBusiness.utilidades.Global utilGlobal = new EndesaBusiness.utilidades.Global();
            cmb_formato_salida.SelectedIndex = 0;
            


            lista_drivers = utilGlobal.GetSystemDriverList();
            string driver = lista_drivers.Find(z => z == "Amazon Redshift (x64)");
            if (driver == null)
            {
                MessageBox.Show("No tiene instalado el driver Amazon RedShift (x64)"
                    + System.Environment.NewLine + "La aplicación no puede continuar y se cerrará.",
                   "ODBC RedShift NO ENCONTRADO",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Information);
                this.Close();
            }

            anio = DateTime.Now.Year;
            mes = (DateTime.Now.Month) - 1;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            fh = new DateTime(anio, mes, dias_del_mes);
            fd = new DateTime(fh.Year, fh.Month, 1);

            //txtFD.Value = fd;
            //txtFH.Value = fh;



        }

        private void btnExcel_Click(object sender, EventArgs e)
        {


            if (lc.Count > 0)
            {
                CurvasCuartoHorarias(lc);
                dgv.Rows.Clear();
                dgv.Refresh();
                lc.Clear();
            }
            else
            {
                lc.Clear();

                //EndesaEntity.medida.PuntoSuministro c = new EndesaEntity.medida.PuntoSuministro();
                //c.cups20 = txtCUPS20.Text;
                //c.fd = Convert.ToDateTime(txt_begin_date.Text);
                //c.fh = Convert.ToDateTime(txt_end_date.Text);
                //lc.Add(c);
                //CurvasCuartoHorarias(lc);
                //dgv.Rows.Clear();
                //dgv.Refresh();
                //lc.Clear();
            }

        }

        private void btnCustomer_Click(object sender, EventArgs e)
        {
            //clientes.Cliente c = new clientes.Cliente();
            //if(txtCNIFDNIC.Text != null && txtCNIFDNIC.Text != "")
            //{
            //    c.fd = Convert.ToDateTime(txt_begin_date.Text);
            //    c.fh = Convert.ToDateTime(txt_end_date.Text);
            //    lblCustomerName.Text = string.Format("{0}", c.Razon_Social(txtCNIFDNIC.Text));

            //    lc = c.lista_cups20;
            //    Carga_DGV(lc);
            //}else
            //{
            //    MessageBox.Show("El campo NIF no puede dejarse en blanco.",
            //    "FrmCurvas.CurvasCuartoHorarias",
            //    MessageBoxButtons.OK,
            //    MessageBoxIcon.Information);
            //    txtCNIFDNIC.Focus();
            //}
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {

            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = false;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    ex = new EndesaBusiness.medida.ExcelCUPS(fileName);
                    if (!ex.hayError)
                    {
                        lc = ex.lista_cups;
                        Carga_DGV(lc);
                    }
                    else
                    {
                        MessageBox.Show(ex.descripcion_error,
                           "Importar Excel",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Information);
                    }

                }
            }
        }

        private void Carga_DGV(List<EndesaEntity.medida.PuntoSuministro> lista)
        {
            lblTotalRegistros.Text = string.Format("Total Registros: {0:#,##0}", lista.Count);
            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void conexiónDataMartToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.FrmCredencialesDatamart f = new FrmCredencialesDatamart();
            //f.Show();
        }

        private void btn_generar_excels_Click(object sender, EventArgs e)
        {
            DialogResult result_1 = MessageBox.Show("¿Desea exportar a Excel los " + string.Format("{0:#,##0}", lc.Count)
                    + " CUPS?"
                    + System.Environment.NewLine
                    + System.Environment.NewLine
                    + @"Se generará un excel por CUPS comprimido excepto para formato ADIF en el directorio c:\Temp",
                    "Exportación de curvas",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result_1 == DialogResult.Yes)
            {
                Cursor.Current = Cursors.WaitCursor;
                if (cmb_formato_salida.SelectedIndex != 2)
                    CurvasCuartoHorarias(lc);
                else
                    CurvasHorarias(lc);
                Cursor.Current = Cursors.Default;
            }
        }


        private void CurvasCuartoHorarias(List<EndesaEntity.medida.PuntoSuministro> lc)
        {

            int f = 0;
            bool firstOnlyOne = true;
            DateTime fechaHora = new DateTime();
            
            //EndesaBusiness.cups.PuntosSuministro ps = new EndesaBusiness.cups.PuntosSuministro();
            //EndesaBusiness.medida.PuntosMedidaPrincipalesVigentes psv = new EndesaBusiness.medida.PuntosMedidaPrincipalesVigentes(lc);
            //EndesaBusiness.medida.FuentesMedida fm = new EndesaBusiness.medida.FuentesMedida();
            //EndesaBusiness.medida.MedidaRedShift redShift = new EndesaBusiness.medida.MedidaRedShift();
            //list_cc = new List<EndesaEntity.medida.CurvaCuartoHorariaInformes>();

            EndesaBusiness.medida.InformeExcelCurvasGestores excelGestores = new EndesaBusiness.medida.InformeExcelCurvasGestores();
            EndesaBusiness.medida.InformeExcelCurvasSCE excelSCE = new EndesaBusiness.medida.InformeExcelCurvasSCE();

            forms.FrmProgressBar pb = new forms.FrmProgressBar();
            double percent = 0;
            int h = 0;

            bool hayError = false;

            EndesaBusiness.medida.Redshift.Curvas curvas = new Curvas();
            EndesaBusiness.medida.Redshift.Estados_Curvas estados_curvas = new Estados_Curvas();

            try
            {
                //// Completamos la lista cups20 con cups13 para indices
                //lc = ps.CompletaCups13(lc);
                //// Completamos la lista con cups15 vigentes
                //lc = psv.CompletaCups15(lc);

                //List<string> lista_cups13 = new List<string>();
                //for (int i = 0; i < lc.Count; i++)
                //    lista_cups13.Add(lc[i].cups13);

                List<string> lista_cups20 = new List<string>();

                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = lc.Count;
                pb.Text = "Exportando... ";

                for (int w = 0; w < lc.Count; w++)
                {

                    //list_cc.Clear();
                    h++;
                    percent = (h / Convert.ToDouble(lc.Count)) * 100;
                    pb.txtDescripcion.Text = "Exportando: " + lc[w].cups20
                        + " desde: " + lc[w].fd.ToString("dd/MM/yyyy")
                        + " hasta: " + lc[w].fh.ToString("dd/MM/yyyy");
                    pb.progressBar.Value = w + 1;
                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    pb.Refresh();

                    if (lc[w].cups20 != null)
                    {

                        lista_cups20.Clear();
                        lista_cups20.Add(lc[w].cups20);
                        curvas.dic_cc.Clear();

                        if (!hayError)
                        {
                            if(radio_facturada.Checked)
                                //hayError = redShift.GetCurvaRedShift(lc[w].cups13, ex.fecha_min, ex.fecha_max, "F");
                                hayError = curvas.GetCurvaGestor(estados_curvas, lista_cups20, lc[w].fd, lc[w].fh, estados_curvas.estados_facturados);
                            else
                                //hayError = redShift.GetCurvaRedShift(lc[w].cups13, ex.fecha_min, ex.fecha_max, "R");
                                hayError = curvas.GetCurvaGestor(estados_curvas, lista_cups20, lc[w].fd, lc[w].fh, estados_curvas.estados_registrados);

                            //hayError = redShift.GetCurvaRedShift2(lc[w].cups20, lc[w].fd, lc[w].fh, "F");

                            //if (redShift.list_cc.Count() != (lc[w].fh - lc[w].fd).TotalDays + 1)
                            //    for (int i = 0; i < lc[w].cups15.Count; i++)
                            //        hayError = redShift.GetCurvaRedShift(lc[w].cups13, ex.fecha_min, ex.fecha_max, "R");


                        }

                        #region SalidaExcel
                        if (!hayError)
                        {
                            //if(redShift.list_cc.Count() > 0)
                            if (curvas.dic_cc.Count() > 0)
                            {
                                if (cmb_formato_salida.SelectedIndex == 0)
                                    //excelGestores.GeneraExcel(lc[w].cups20, redShift.list_cc);
                                    excelGestores.GeneraExcel(lc[w].cups20, curvas.dic_cc);
                                else
                                {
                                    //excelSCE.GeneraExcel(lc[w].cups20, redShift.list_cc);
                                }

                            }
                            else
                            {
                                MessageBox.Show("No hay datos para la curva "
                                    + lc[w].cups20 
                                    + Environment.NewLine
                                    + " del " + lc[w].fd.ToString("dd/MM/yyyy")
                                    + " hasta " + lc[w].fh.ToString("dd/MM/yyyy"),
                                      "Exportación Curvas RedShift",
                                      MessageBoxButtons.OK,
                                      MessageBoxIcon.Information);
                            }

                            





                            //f = 0;
                            //firstOnlyOne = true;

                            //FileInfo file = new FileInfo(@"c:\Temp" + @"\" + redShift.list_cc[0].cups22.Substring(0, 20)
                            //    + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
                            //ExcelPackage excelPackage = new ExcelPackage(file);
                            //var workSheet = excelPackage.Workbook.Worksheets.Add("CC_1");
                            //var headerCells = workSheet.Cells[1, 1, 1, 25];
                            //var headerFont = headerCells.Style.Font;

                            //list_cc = redShift.list_cc.OrderBy(z => z.cups15).ThenBy(z => z.fecha).ToList();

                            //for (int x = 0; x < list_cc.Count; x++)
                            //{
                            //    fechaHora = list_cc[x].fecha;

                            //    #region Cabecera
                            //    if (firstOnlyOne)
                            //    {
                            //        f++;
                            //        workSheet.Cells[f, 1].Value = "CUPS15";
                            //        workSheet.Cells[f, 2].Value = "CUPS22";
                            //        workSheet.Cells[f, 3].Value = "FECHA";
                            //        workSheet.Cells[f, 4].Value = "HORA";
                            //        workSheet.Cells[f, 5].Value = "Energía Activa Horaria (kWh)";
                            //        workSheet.Cells[f, 6].Value = "Energía Reactiva Horaria (kVar)";
                            //        workSheet.Cells[f, 7].Value = "Potencia Activa";
                            //        workSheet.Cells[f, 8].Value = "Cuarto Horaria Activa";
                            //        workSheet.Cells[f, 9].Value = "FUENTE FINAL";
                            //        workSheet.Cells[f, 10].Value = "ESTADO";

                            //        firstOnlyOne = false;

                            //    }
                            //    #endregion

                            //    for (int p = 1; p <= list_cc[x].numPeriodos; p++)
                            //    {
                            //        f++;
                            //        #region 23 Periodos        
                            //        if (list_cc[x].numPeriodos == 23 && p > 2)
                            //        {
                            //            if (p == 3)
                            //                fechaHora = fechaHora.AddHours(1);

                            //            workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                            //            workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                            //            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            //            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            //            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            //            workSheet.Cells[f, 5].Value = list_cc[x].a[p + 1];
                            //            workSheet.Cells[f, 5].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 6].Value = list_cc[x].r[p + 1];
                            //            workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 7].Value = list_cc[x].value[((p + 1) * 4) - 3];
                            //            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4) - 3] / 4;
                            //            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            //            if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                            //                list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                            //            {
                            //                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                            //                    Convert.ToInt32(list_cc[x].fa[p + 1]));
                            //            }

                            //            workSheet.Cells[f, 10].Value = list_cc[x].estado;

                            //            fechaHora = fechaHora.AddMinutes(15);
                            //            f++;
                            //            workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                            //            workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                            //            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            //            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            //            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            //            workSheet.Cells[f, 7].Value = list_cc[x].value[((p + 1) * 4) - 2];
                            //            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4) - 2] / 4;
                            //            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            //            if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                            //                list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                            //            {
                            //                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                            //                    Convert.ToInt32(list_cc[x].fa[p + 1]));
                            //            }

                            //            workSheet.Cells[f, 10].Value = list_cc[x].estado;

                            //            fechaHora = fechaHora.AddMinutes(15);
                            //            f++;
                            //            workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                            //            workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                            //            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            //            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            //            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            //            workSheet.Cells[f, 7].Value = list_cc[x].value[((p + 1) * 4) - 1];
                            //            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4) - 1] / 4;
                            //            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            //            if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                            //                list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                            //            {
                            //                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                            //                    Convert.ToInt32(list_cc[x].fa[p + 1]));
                            //            }

                            //            workSheet.Cells[f, 10].Value = list_cc[x].estado;

                            //            fechaHora = fechaHora.AddMinutes(15);
                            //            f++;
                            //            workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                            //            workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                            //            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            //            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            //            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            //            workSheet.Cells[f, 7].Value = list_cc[x].value[((p + 1) * 4)];
                            //            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 8].Value = list_cc[x].value[((p + 1) * 4)] / 4;
                            //            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                            //            fechaHora = fechaHora.AddMinutes(15);

                            //            if (list_cc[x].fc[p + 1] != null && list_cc[x].fc[p + 1] != "" &&
                            //                list_cc[x].fa[p + 1] != null && list_cc[x].fa[p + 1] != "")
                            //            {
                            //                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p + 1]),
                            //                    Convert.ToInt32(list_cc[x].fa[p + 1]));
                            //            }

                            //            workSheet.Cells[f, 10].Value = list_cc[x].estado;
                            //        }
                            //        #endregion
                            //        #region 25 Periodos
                            //        else if (list_cc[x].numPeriodos == 25 && p > 2)
                            //        {
                            //            if (p == 4)
                            //                fechaHora = fechaHora.AddHours(-1);

                            //            workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                            //            workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                            //            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            //            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            //            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            //            workSheet.Cells[f, 5].Value = list_cc[x].a[p];
                            //            workSheet.Cells[f, 5].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 6].Value = list_cc[x].r[p];
                            //            workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 3];
                            //            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 3] / 4;
                            //            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            //            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            //                (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            //            {

                            //                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                            //                    Convert.ToInt32(list_cc[x].fa[p]));
                            //            }

                            //            workSheet.Cells[f, 10].Value = list_cc[x].estado;

                            //            fechaHora = fechaHora.AddMinutes(15);
                            //            f++;
                            //            workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                            //            workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                            //            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            //            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            //            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            //            workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 2];
                            //            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 2] / 4;
                            //            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            //            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            //                 (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            //            {

                            //                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                            //                    Convert.ToInt32(list_cc[x].fa[p]));
                            //            }

                            //            workSheet.Cells[f, 10].Value = list_cc[x].estado;

                            //            fechaHora = fechaHora.AddMinutes(15);
                            //            f++;
                            //            workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                            //            workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                            //            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            //            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            //            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            //            workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 1];
                            //            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 1] / 4;
                            //            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            //            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            //                 (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            //            {

                            //                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                            //                    Convert.ToInt32(list_cc[x].fa[p]));
                            //            }

                            //            workSheet.Cells[f, 10].Value = list_cc[x].estado;

                            //            fechaHora = fechaHora.AddMinutes(15);
                            //            f++;
                            //            workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                            //            workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                            //            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            //            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            //            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            //            workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4)];
                            //            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4)] / 4;
                            //            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                            //            fechaHora = fechaHora.AddMinutes(15);

                            //            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            //                (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            //            {
                            //                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]), Convert.ToInt32(list_cc[x].fa[p]));
                            //            }

                            //            workSheet.Cells[f, 10].Value = list_cc[x].estado;

                            //        }
                            //        #endregion
                            //        #region 24 Periodos
                            //        else
                            //        {
                            //            workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                            //            workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                            //            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            //            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            //            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            //            workSheet.Cells[f, 5].Value = list_cc[x].a[p];
                            //            workSheet.Cells[f, 5].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 6].Value = list_cc[x].r[p];
                            //            workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 3];
                            //            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 3] / 4;
                            //            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            //            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            //                (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            //            {
                            //                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                            //                    Convert.ToInt32(list_cc[x].fa[p]));
                            //            }

                            //            workSheet.Cells[f, 10].Value = list_cc[x].estado;

                            //            fechaHora = fechaHora.AddMinutes(15);
                            //            f++;
                            //            workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                            //            workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                            //            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            //            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            //            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            //            workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 2];
                            //            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 2] / 4;
                            //            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";

                            //            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            //                (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            //            {
                            //                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                            //                    Convert.ToInt32(list_cc[x].fa[p]));
                            //            }

                            //            workSheet.Cells[f, 10].Value = list_cc[x].estado;

                            //            fechaHora = fechaHora.AddMinutes(15);
                            //            f++;
                            //            workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                            //            workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                            //            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            //            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            //            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            //            workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4) - 1];
                            //            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4) - 1] / 4;
                            //            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";


                            //            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            //                (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            //            {
                            //                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                            //                    Convert.ToInt32(list_cc[x].fa[p]));
                            //            }

                            //            workSheet.Cells[f, 10].Value = list_cc[x].estado;

                            //            fechaHora = fechaHora.AddMinutes(15);
                            //            f++;
                            //            workSheet.Cells[f, 1].Value = list_cc[x].cups15;
                            //            workSheet.Cells[f, 2].Value = list_cc[x].cups22;
                            //            workSheet.Cells[f, 3].Value = fechaHora.Date;
                            //            workSheet.Cells[f, 3].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            //            workSheet.Cells[f, 4].Value = fechaHora.ToString("HH:mm:ss");
                            //            workSheet.Cells[f, 7].Value = list_cc[x].value[(p * 4)];
                            //            workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                            //            workSheet.Cells[f, 8].Value = list_cc[x].value[(p * 4)] / 4;
                            //            workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                            //            fechaHora = fechaHora.AddMinutes(15);

                            //            if ((list_cc[x].fc[p] != null && list_cc[x].fc[p] != "") &&
                            //                (list_cc[x].fa[p] != null && list_cc[x].fa[p] != ""))
                            //            {
                            //                workSheet.Cells[f, 9].Value = fm.FuenteFinal(Convert.ToInt32(list_cc[x].fc[p]),
                            //                    Convert.ToInt32(list_cc[x].fa[p]));
                            //            }

                            //            workSheet.Cells[f, 10].Value = list_cc[x].estado;
                            //        }

                            //        #endregion
                            //    }
                            //}

                            //var allCells = workSheet.Cells[1, 1, f, 11];
                            //allCells.AutoFitColumns();
                            //excelPackage.Save();

                            //Console.WriteLine("Comprimiendo archivo ...");
                            //zip.ComprimirArchivo(file.FullName, file.FullName.Replace(".xlsx", ".zip"));                            
                            //file.Delete();


                        }
                        #endregion

                        
                    }

                }

                pb.Close();
                MessageBox.Show("Exportación Finalizada", "Exportación de Curvas", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //else if (!hayError)// Caso en el que no se han encontrado datos
                //{
                //    // this.RespondeCorreo(mail, "No existen curvas facturadas para los cups y periodos solicitados.");
                //    // this.RespondeMail(mail, TipoRespuestaCorreo.SinCurvas);

                //}

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "CurvasCuartoHorarias", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CurvasHorarias(List<EndesaEntity.medida.PuntoSuministro> lc)
        {
            int f = 1;
            int c = 1;
            DateTime fd = new DateTime();
            DateTime fh = new DateTime();
            DateTime fechaHora = new DateTime();

            EndesaBusiness.medida.Redshift.Curvas cc_bi;
            EndesaBusiness.medida.Redshift.Estados_Curvas estados_curvas;
            Dictionary<string, string> dic_cups20_adif = new Dictionary<string, string>();


            fd = new DateTime(4999,12,31);
            fh = DateTime.MinValue;

            foreach(EndesaEntity.medida.PuntoSuministro p in lc)
            {
                string v_cups20;
                if (!dic_cups20_adif.TryGetValue(p.cups20, out v_cups20))
                    dic_cups20_adif.Add(p.cups20, p.cups20);

                if(fd > p.fd)
                    fd = p.fd;

                if(fh < p.fh)
                    fh = p.fh;

            }

            estados_curvas = new Estados_Curvas();
            cc_bi = new EndesaBusiness.medida.Redshift.Curvas();
            cc_bi.CargaMedidaHoraria(estados_curvas, dic_cups20_adif.Select(z => z.Key).ToList(), fd, fh, estados_curvas.estados_facturados);
            cc_bi.CargaMedidaHoraria(estados_curvas, dic_cups20_adif.Select(z => z.Key).ToList(), fd, fh, estados_curvas.estados_registrados);

            FileInfo file = new FileInfo(@"c:\Temp" + @"\" + "ADIF"
               + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(file);
            var workSheet = excelPackage.Workbook.Worksheets.Add("CH");

            var headerCells = workSheet.Cells[1, 1, 1, 26];
            var headerFont = headerCells.Style.Font;


            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "FECHA"; c++;
            workSheet.Cells[f, c].Value = "HORA"; c++;
            workSheet.Cells[f, c].Value = "AE"; c++;
            workSheet.Cells[f, c].Value = "AS"; c++;
            workSheet.Cells[f, c].Value = "R1"; c++;
            workSheet.Cells[f, c].Value = "R2"; c++;
            workSheet.Cells[f, c].Value = "R3"; c++;
            workSheet.Cells[f, c].Value = "R4"; c++;
            workSheet.Cells[f, c].Value = "CUPS22"; c++;


            foreach(KeyValuePair< string, List<CurvaCuartoHoraria>> p in cc_bi.dic_cc)
            {

                

                foreach (CurvaCuartoHoraria pp in p.Value)
                {

                    fechaHora = pp.fecha;
                    for (int i = 1; i <= 24; i++)
                    {
                        f++;
                        workSheet.Cells[f, 1].Value = pp.fecha;
                        workSheet.Cells[f, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, 2].Value = fechaHora.ToString("HH:mm");
                        workSheet.Cells[f, 3].Value = pp.a[i];
                        workSheet.Cells[f, 3].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 4].Value = pp.r[i];
                        workSheet.Cells[f, 4].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 5].Value = pp.r1[i];
                        workSheet.Cells[f, 5].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 6].Value = pp.r2[i];
                        workSheet.Cells[f, 6].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 7].Value = pp.r3[i];
                        workSheet.Cells[f, 7].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 8].Value = pp.r4[i];
                        workSheet.Cells[f, 8].Style.Numberformat.Format = "#,##0";
                        workSheet.Cells[f, 9].Value = pp.cups22;

                        fechaHora = fechaHora.AddHours(1);
                    }
                    


                }
            }


            var allCells = workSheet.Cells[1, 1, f, 9];
            allCells.AutoFitColumns();
            excelPackage.Save();

            MessageBox.Show("Exportación Finalizada", "Exportación de Curvas", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
       

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        

        private void btnImportExcel_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Excel files|*.xlsx";
            d.Multiselect = false;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    ex = new EndesaBusiness.medida.ExcelCUPS(fileName);
                    if (!ex.hayError)
                    {
                        lc = ex.lista_cups;
                        Carga_DGV(lc);
                    }
                    else
                    {
                        MessageBox.Show(ex.descripcion_error,
                           "Error en el formato del fichero",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Information);
                    }

                }
            }
        }

       

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            //List<string> lista_cups = new List<string>();
            //lista_cups.Add(txtCUPS20.Text);
            //EndesaBusiness.medida.CCRD cc = new EndesaBusiness.medida.CCRD(lista_cups, txtFD.Value, txtFH.Value, "F");
            //dgv_cc.AutoGenerateColumns = true;
            //dgv_cc.DataSource = cc.lista;
        } 

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void FrmCurvas_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Medida", "FrmCurvas", "N/A");
        }
    }
}
