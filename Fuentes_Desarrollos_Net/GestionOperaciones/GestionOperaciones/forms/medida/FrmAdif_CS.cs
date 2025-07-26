using EndesaBusiness.medida.Redshift;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.medida
{
    public partial class FrmAdif_CS : Form
    {
        EndesaBusiness.adif.ProcesosFunciones pf = new EndesaBusiness.adif.ProcesosFunciones();
        EndesaBusiness.adif.InventarioFunciones invf;
        EndesaBusiness.adif.MedidaAdif med;
        List<EndesaEntity.medida.AdifInformeMedida> informe;
        EndesaBusiness.adif.NuevaMedidaADIF medida_adif;
        // 20181126 GO.medida.curvasdecarga.CurvaCuartoHoraria cc_sce;
        EndesaBusiness.medida.CurvaCuartoHorariaSCE cc_sce;

        EndesaBusiness.medida.PuntosMedidaPrincipalesVigentes pm;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        EndesaBusiness.medida.Redshift.Estados_Curvas estados_curvas;
        EndesaBusiness.medida.Redshift.Curvas cc_bi;

        EndesaBusiness.adif.CierresEnergia cierres_energia;

        EndesaBusiness.utilidades.Param p;

        public FrmAdif_CS()
        {
            int anio;
            int mes;
            int dias_del_mes;
            DateTime fh = new DateTime();
            DateTime fd = new DateTime();
            DateTime mesAnterior = new DateTime();

            usage.Start("Medida", "FrmAdif_CS" ,"N/A");

            p = new EndesaBusiness.utilidades.Param("adif_param", EndesaBusiness.servidores.MySQLDB.Esquemas.MED);

            InitializeComponent();

            mesAnterior = DateTime.Now.AddMonths(-1);
            anio = mesAnterior.Year;
            mes = mesAnterior.Month;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);
                        
            fh = new DateTime(anio, mes, dias_del_mes);
            fd = new DateTime(fh.Year, fh.Month, 1);

            txtFFACTDES.Value = fd;
            txtFFACTHAS.Value = fh;

            chkDiff.CheckState = CheckState.Unchecked;
            InitListBox(Convert.ToDateTime(txtFFACTDES.Value), Convert.ToDateTime(txtFFACTHAS.Value));

            estados_curvas = new Estados_Curvas();

        }

        private void InitListBox(DateTime fd, DateTime fh)
        {
            EndesaBusiness.adif.AdifLotes lotes = new EndesaBusiness.adif.AdifLotes(fd, fh);

            try
            {
                for (int i = listBoxLotes.Items.Count - 1; i >= 0; i--)
                {
                    listBoxLotes.Items.RemoveAt(i);
                }              
                
                for(int i = 0; i < lotes.lista_lotes.Count; i++)                
                    listBoxLotes.Items.Add(lotes.lista_lotes[i]);
                
                
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "InitListBox",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            }
            


        }

        private void FrmAdif_CS_Load(object sender, EventArgs e)
        {
            List<string> lista_drivers = new List<string>();

            EndesaBusiness.utilidades.Global utilGlobal = new EndesaBusiness.utilidades.Global();
            lista_drivers = utilGlobal.GetSystemDriverList();
            string driver = lista_drivers.Find(z => z.Contains("Redshift"));
            if (driver == null)
            {
                MessageBox.Show("No tiene instalado el driver Amazon RedShift (x64)"
                    + System.Environment.NewLine + "La aplicación no puede continuar y se cerrará.",
                   "ODBC RedShift NO ENCONTRADO",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Information);
                this.Close();
            }

            lbl_fecha_importacion.Text = string.Format("Última importacion medida proporcionada por ADIF: {0}", pf.GetLastProcess("Importar PO1011").ToString("dd/MM/yyyy"));            

            cierres_energia = new EndesaBusiness.adif.CierresEnergia();

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            List<EndesaEntity.medida.AdifInformeMedida> informe = new List<EndesaEntity.medida.AdifInformeMedida>();
            List<string> lista_cups20_adif = new List<string>();
            List<string> lista_cups13_adif = new List<string>();

            Dictionary<string, string> dic_cups20_adif = new Dictionary<string, string>();
            Dictionary<string, string> dic_cups13_adif = new Dictionary<string, string>();

            string cups20 = null;
            
            List<string> lista_lotes = new List<string>();
            
            DateTime fd;
            DateTime fh;

            double calculo;
            double denominador;
            double umbral;
            Boolean tiene_curvas;

            Cursor.Current = Cursors.WaitCursor;

            if (txtcups20.Text != "")
                cups20 = txtcups20.Text;
            
                        
            fd = Convert.ToDateTime(txtFFACTDES.Text);
            fh = Convert.ToDateTime(txtFFACTHAS.Text);

            if (listBoxLotes.SelectedItems.Count > 0)
                foreach (Object values in listBoxLotes.SelectedItems)
                    lista_lotes.Add(values.ToString());

            EndesaBusiness.adif.Param pp = new EndesaBusiness.adif.Param();
            umbral = Convert.ToDouble(pp.GetParam("Umbral_diferencia"));

            Cursor.Current = Cursors.WaitCursor;

            invf = new EndesaBusiness.adif.InventarioFunciones(fd, fh, cups20, lista_lotes);
            // med = new GO.medida.adif.MedidaAdif(fd, fh, invf.dic_inventario);            



            foreach (KeyValuePair<string, EndesaEntity.medida.AdifInventario> p in invf.dic_inventario)
            {
                //202006 lista_cups20_adif.Add(p.Value.cups20);
                //202006 lista_cups13_adif.Add(p.Value.cups13);

                string v_cups20;
                if (!dic_cups20_adif.TryGetValue(p.Value.cups20, out v_cups20))
                    dic_cups20_adif.Add(p.Value.cups20, p.Value.cups20);

                //string v_cups13;
                //if (!dic_cups13_adif.TryGetValue(p.Value.cups13, out v_cups13))
                //    dic_cups13_adif.Add(p.Value.cups13, p.Value.cups20);

            }

            // MEDIDA ADIF
            // ===========
            //medida_adif = new EndesaBusiness.adif.NuevaMedidaADIF(dic_cups20_adif.Select(z => z.Key).ToList(), fd, fh);
            medida_adif = new EndesaBusiness.adif.NuevaMedidaADIF();
            medida_adif.CargaMedidaHoraria(dic_cups20_adif.Select(z => z.Key).ToList(), fd, fh);
            
            
            // PUNTOS DE MEDIDA            
            pm = new EndesaBusiness.medida.PuntosMedidaPrincipalesVigentes(dic_cups20_adif.Select(z => z.Key).ToList(), fd, fh);

            // Curvas Resumen
            // ===============                   

            //EndesaBusiness.medida.CurvaResumenBI cr_bi = new EndesaBusiness.medida.CurvaResumenBI(dic_cups20_adif.Select(z => z.Key).ToList(), fd, fh);
            EndesaBusiness.medida.CurvaResumenBI cr_bi = new EndesaBusiness.medida.CurvaResumenBI(pm.lista_cups22, fd, fh);

            // Curvas CuartoHorarias

            cc_bi = new EndesaBusiness.medida.Redshift.Curvas();
            // cc_bi.CargaMedidaCuartoHoraria(estados_curvas, dic_cups20_adif.Select(z => z.Key).ToList(), fd, fh, estados_curvas.estados_facturados);
            // cc_bi.CargaMedidaCuartoHoraria(estados_curvas, dic_cups20_adif.Select(z => z.Key).ToList(), fd, fh, estados_curvas.estados_registrados);

            cc_bi.CargaMedidaCuartoHoraria(estados_curvas, dic_cups20_adif.Select(z => z.Key).ToList(), fd, fh, estados_curvas.estados_facturados);
            cc_bi.CargaMedidaHorariaCups22(estados_curvas, pm.lista_cups22, fd, fh, estados_curvas.estados_registrados);

            informe = new List<EndesaEntity.medida.AdifInformeMedida>();
            foreach (KeyValuePair<string, EndesaEntity.medida.AdifInventario> p in invf.dic_inventario)
            {
                //Variable auxiliar para marcar si existe curvas en algún punto de suministro, ya que si p.e. tiene dos puntos y el último no tiene curvas
                // en línea 267 aprox. no entraría
                tiene_curvas = false;
                p.Value.activa_entrante = medida_adif.TotalActivaEntrante(p.Value.cups20, p.Value.ffactdes, p.Value.ffacthas);
                p.Value.activa_saliente = medida_adif.TotalActivaSaliente(p.Value.cups20, p.Value.ffactdes, p.Value.ffacthas);

                if (!medida_adif.existe_curva && !tiene_curvas)
                {
                    p.Value.activa_entrante = 0;
                    p.Value.activa_saliente = 0;
                }

                if (p.Value.devolucion_de_energia)
                    p.Value.neteada = medida_adif.Neteo(p.Value.cups20, p.Value.ffactdes, p.Value.ffacthas);

                
                p.Value.cierres_energia = cierres_energia.ExisteCierre(p.Value.cups20, p.Value.ffactdes, p.Value.ffacthas);


                EndesaEntity.medida.PuntoSuministro punto_suministro;
                if(pm.dic.TryGetValue(p.Value.cups20,out punto_suministro))
                {
                    for(int i = 0; i < punto_suministro.cups22.Count; i++)
                    {
                        cr_bi.GetCurva(punto_suministro.cups22[i], p.Value.ffactdes, p.Value.ffacthas);

                        if (cr_bi.existe_curva)
                        {
                            tiene_curvas = true;
                            p.Value.resumen_sce_a += cr_bi.activa;
                            p.Value.resumen_sce_r += cr_bi.reactiva;
                            p.Value.estado_curva = cr_bi.estado;
                            p.Value.fuente = cr_bi.fuente;
                            p.Value.dias = cr_bi.dias;                            

                        }
                        
                        if(cc_bi.ExisteCurva(punto_suministro.cups22[i], p.Value.ffactdes, p.Value.ffacthas))
                        {
                            p.Value.sce_a += Convert.ToDouble(cc_bi.TotalActiva(punto_suministro.cups22[i], p.Value.ffactdes, p.Value.ffacthas));
                            p.Value.sce_r = Convert.ToDouble(cc_bi.TotalReactiva(punto_suministro.cups22[i], p.Value.ffactdes, p.Value.ffacthas));
                        }

                        



                    }

                   

                }


                EndesaEntity.medida.AdifInformeMedida c = new EndesaEntity.medida.AdifInformeMedida();
                //if (cr_bi.existe_curva) 18/06/2024 GUS: añadimos variable control si existe curva en algún punto de medida, no solo en el último
                if(tiene_curvas)
                {
                    c.resumen_sce_a = p.Value.resumen_sce_a;
                    c.resumen_sce_r = p.Value.resumen_sce_r;
                    c.fuente = p.Value.fuente;
                    c.dias = p.Value.dias;
                    
                    c.sce_a = p.Value.sce_a;
                    c.sce_r = p.Value.sce_r;
                }

                c.estado_curva = p.Value.estado_curva;

                c.adif_a = p.Value.adif_a;
                //c.adif_r = p.Value.adif_r;
                c.codigo = p.Value.codigo;
                c.comentarios = p.Value.comentarios;                
                //c.cups13 = p.Value.cups13;
                c.cups20 = p.Value.cups20;

                c.perdidas = p.Value.perdidas;
                c.valor_kvas = p.Value.valor_kvas;

                c.medida_en_baja = p.Value.medida_en_baja;
                c.devolucion_de_energia = p.Value.devolucion_de_energia;
                c.cierres_energia = p.Value.cierres_energia;

                c.activa_entrante = p.Value.activa_entrante;
                c.activa_saliente = p.Value.activa_saliente;
                c.neteada = p.Value.neteada;

                if (p.Value.devolucion_de_energia)
                    c.neteada = p.Value.neteada;

                c.estado_contrato = p.Value.estado_contrato;                

                c.tarifa = p.Value.tarifa;
                c.lote = p.Value.lote;
                c.ffactdes = p.Value.ffactdes;
                c.ffacthas = p.Value.ffacthas;

                if (p.Value.devolucion_de_energia)                
                    p.Value.dif_sce_adif_a = p.Value.sce_a - p.Value.neteada;                    
                else                
                    p.Value.dif_sce_adif_a = p.Value.sce_a - p.Value.activa_entrante;
                    
                
                p.Value.dif_sce_adif_r = p.Value.sce_r - p.Value.adif_r;

                if (p.Value.sce_a == 0)
                    denominador = 1;
                else
                    denominador = p.Value.sce_a;


                if (p.Value.devolucion_de_energia)
                    calculo = (Math.Abs(p.Value.sce_a - p.Value.neteada) / denominador) * 100;
                else
                    calculo = (Math.Abs(p.Value.sce_a - p.Value.activa_entrante) / denominador) * 100;

                if (p.Value.devolucion_de_energia)
                {
                    if (p.Value.sce_a == p.Value.neteada)
                    {
                        p.Value.resultado = "BI=ADIF";
                        p.Value.enviado_facturar = "OK";
                    }
                    else if (umbral > calculo || Math.Abs(p.Value.sce_a - p.Value.neteada) < 100)
                    {
                        p.Value.resultado = "%ERROR ASUMIBLE";
                        p.Value.enviado_facturar = "OK";
                    }
                    else
                    {
                        p.Value.resultado = "REVISAR";
                        p.Value.enviado_facturar = "REVISAR";
                    }
                }
                else
                {
                    if (p.Value.sce_a == p.Value.activa_entrante)
                    {
                        p.Value.resultado = "BI=ADIF";
                        p.Value.enviado_facturar = "OK";
                    }
                    else if (umbral > calculo || Math.Abs(p.Value.sce_a - p.Value.activa_entrante) < 100)
                    {
                        p.Value.resultado = "%ERROR ASUMIBLE";
                        p.Value.enviado_facturar = "OK";
                    }
                    else
                    {
                        p.Value.resultado = "REVISAR";
                        p.Value.enviado_facturar = "REVISAR";
                    }
                }

                

                c.dif_sce_adif_a = p.Value.dif_sce_adif_a;
                c.dif_sce_adif_r = p.Value.dif_sce_adif_r;
                c.resultado = p.Value.resultado;
                c.enviado_facturar = p.Value.enviado_facturar;

                informe.Add(c);

            }

            this.dgv.AutoGenerateColumns = false;
            if (informe.Count > 0)
            {
                if (chkDiff.Checked)
                {
                    List<EndesaEntity.medida.AdifInformeMedida> lista = informe.Where(kvp =>  kvp.dif_sce_adif_a != 0).ToList();
                    lblTotalRegistros.Text = string.Format("Total Registros: {0:#,##0}", lista.Count());
                    this.dgv.DataSource = lista;
                    // GUS: 16/06/2024
                    // MARCAR FONDO ROJO Si
                    //      [ADIF Entrante] (12) != [BI Activa] (15)
                    //      Y
                    //      NOT ([Devolución de energía] (6) == true AND ([Neteo] (14) == [BI Activa] (15))
                    foreach (DataGridViewRow row in dgv.Rows)
                    {
                        //if (Convert.ToInt32(row.Cells[12].Value) != Convert.ToInt32(row.Cells[15].Value))
                        if ((Convert.ToInt32(row.Cells[12].Value) != Convert.ToInt32(row.Cells[15].Value)) && !(Convert.ToBoolean(row.Cells[6].Value) && (Convert.ToInt32(row.Cells[14].Value) == Convert.ToInt32(row.Cells[15].Value))))
                            row.DefaultCellStyle.BackColor = Color.LightPink;
                    }  

                }
                else
                {
                    lblTotalRegistros.Text = string.Format("Total Registros: {0:#,##0}", informe.Count());
                    this.dgv.DataSource = informe;
                    // GUS: 16/06/2024
                    // MARCAR FONDO ROJO Si
                    //      [ADIF Entrante] (12) != [BI Activa] (15)
                    //      Y
                    //      NOT ([Devolución de energía] (6) == true AND ([Neteo] (14) == [BI Activa] (15))
                    foreach (DataGridViewRow row in dgv.Rows)
                    {
                        //if (Convert.ToInt32(row.Cells[12].Value) != Convert.ToInt32(row.Cells[15].Value))
                        if ((Convert.ToInt32(row.Cells[12].Value) != Convert.ToInt32(row.Cells[15].Value)) && !(Convert.ToBoolean(row.Cells[6].Value) && (Convert.ToInt32(row.Cells[14].Value) == Convert.ToInt32(row.Cells[15].Value))))
                            row.DefaultCellStyle.BackColor = Color.LightPink;

                        //Console.WriteLine("Valor columna [Cierres energía]: " + Convert.ToBoolean(row.Cells[6].Value));
                    }

                }
            }else
            {
                dgv.DataSource = null;

                MessageBox.Show("Sin datos",
              "Informe terminado",
              MessageBoxButtons.OK,
              MessageBoxIcon.Information);
            }
            
                                    

            Cursor.Current = Cursors.Default;
        }

       

        private void ExportExcel(string fichero)
        {
            int f = 1;
            int c = 1;

            FileInfo fileInfo = new FileInfo(fichero);

            if (fileInfo.Exists)
                fileInfo.Delete();


            ExcelPackage excelPackage;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            excelPackage = new ExcelPackage(fileInfo);
            //ExcelPackage excelPackage = new ExcelPackage();
            var workSheet = excelPackage.Workbook.Worksheets.Add("ADIF");



            var headerCells = workSheet.Cells[1, 1, 1, 27];
            var headerFont = headerCells.Style.Font;

           
            //headerFont.SetFromFont(new Font("Times New Roman", 12)); //Do this first
            headerFont.Bold = true;
            //headerFont.Italic = true;
            
            workSheet.Cells[f, c].Value = "CUPSREE"; c++;
            workSheet.Cells[f, c].Value = "Tarifa"; c++;
            workSheet.Cells[f, c].Value = "Lote"; c++;
            workSheet.Cells[f, c].Value = "FFACTDES"; c++;
            workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
            workSheet.Cells[f, c].Value = "Medida en baja"; c++;
            workSheet.Cells[f, c].Value = "Devolución de energía"; c++;
            workSheet.Cells[f, c].Value = "Cierres energía"; c++;
            workSheet.Cells[f, c].Value = "DIAS"; c++;
            workSheet.Cells[f, c].Value = "FUENTE"; c++;
            workSheet.Cells[f, c].Value = "ESTADO"; c++;
            workSheet.Cells[f, c].Value = "Resumen BI Activa"; c++;
            workSheet.Cells[f, c].Value = "ADIF Entrante"; c++;
            workSheet.Cells[f, c].Value = "ADIF Saliente"; c++;
            workSheet.Cells[f, c].Value = "Neteada"; c++;
            workSheet.Cells[f, c].Value = "BI Activa"; c++;
            workSheet.Cells[f, c].Value = "REE Activa"; c++;
            workSheet.Cells[f, c].Value = "DIF BI-ADIF E."; c++;
            workSheet.Cells[f, c].Value = "Estado BI"; c++;
            workSheet.Cells[f, c].Value = "RESULTADO"; c++;            
            workSheet.Cells[f, c].Value = "ENVIADO A FACTURAR"; c++;
            workSheet.Cells[f, c].Value = "Valor (KVA)"; c++;
            workSheet.Cells[f, c].Value = "Pérdidas"; c++;



            // foreach (KeyValuePair<string, EndesaEntity.medida.AdifInformeMedida> p in !chkDiff.Checked ? med.dic_informe : med.dic_informe.Where(kvp => kvp.Value.dif_sce_adif_a != 0))
            foreach (KeyValuePair<string, EndesaEntity.medida.AdifInventario> p in !chkDiff.Checked ? invf.dic_inventario : invf.dic_inventario.Where(kvp => kvp.Value.dif_sce_adif_a != 0))
            {
                f++;
                c = 1;                
                workSheet.Cells[f, c].Value = p.Value.cups20; c++;
                workSheet.Cells[f, c].Value = p.Value.tarifa; c++;
                workSheet.Cells[f, c].Value = p.Value.lote; c++;
                workSheet.Cells[f, c].Value = p.Value.ffactdes; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.Value.ffacthas; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                if (p.Value.medida_en_baja)
                    workSheet.Cells[f, c].Value = "Sí";
                else
                    workSheet.Cells[f, c].Value = "No";

                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                if (p.Value.devolucion_de_energia)
                    workSheet.Cells[f, c].Value = "Sí";
                else
                    workSheet.Cells[f, c].Value = "No";

                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;

                if (p.Value.cierres_energia)
                    workSheet.Cells[f, c].Value = "Sí";
                else
                    workSheet.Cells[f, c].Value = "No";

                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;


                workSheet.Cells[f, c].Value = p.Value.dias; c++;
                workSheet.Cells[f, c].Value = p.Value.fuente; c++;
                workSheet.Cells[f, c].Value = p.Value.estado_curva; c++;

                workSheet.Cells[f, c].Value = p.Value.resumen_sce_a; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.activa_entrante; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.activa_saliente; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.neteada; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.sce_a; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.ree_a; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.dif_sce_adif_a; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;


                workSheet.Cells[f, c].Value = p.Value.estado_contrato; c++;
                workSheet.Cells[f, c].Value = p.Value.resultado; c++;
                workSheet.Cells[f, c].Value = p.Value.enviado_facturar; c++;
                workSheet.Cells[f, c].Value = p.Value.valor_kvas; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = p.Value.perdidas; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

            }


            var allCells = workSheet.Cells[1, 1, f, c];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:W1"].AutoFilter = true;
            allCells.AutoFitColumns();
            excelPackage.Save();
        

        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {

            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del archivo Estado medida ADIF";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                this.ExportExcel(save.FileName);
                MessageBox.Show("Informe terminado.",
                  "Exportación a Excel",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Information);
                
                System.Diagnostics.Process.Start(save.FileName);
                
            }
        }

        private void btnCurvas_Click(object sender, EventArgs e)
        {

        }

        private void crearComparativaCurvaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool exportado_algo = false;
            List<string> lista_cups20 = new List<string>();
            string fichero = "";
            DateTime ahora = DateTime.Now;
            bool firstOnly = true;
            string ruta = "";
            bool continua = false;

            

            foreach (DataGridViewRow row in dgv.SelectedRows)
            {

                if (firstOnly)
                {

                    fichero = "ARCHIVO_COMPARATIVA_DE_CURVA_" + row.Cells[0].Value.ToString() + "_"
                        + Convert.ToDateTime(row.Cells[3].Value).ToString("yyyyMM") + "_"
                    + ahora.ToString("yyyyMMdd_HHmmss") + ".xlsx";

                    SaveFileDialog save = new SaveFileDialog();
                    save.Title = "Ubicación del archivo de comparativa de curva";
                    save.FileName = fichero;
                    save.AddExtension = true;
                    save.DefaultExt = "xlsx";
                    save.Filter = "Ficheros xslx (*.xlsx)|*.*";
                    DialogResult result = save.ShowDialog();


                   
                
                    //fichero = ruta + "\\" + "ARCHIVO_COMPARATIVA_DE_CURVA_" + row.Cells[0].Value.ToString() + "_"
                    //    + Convert.ToDateTime(row.Cells[3].Value).ToString("yyyyMM") + "_"
                    //+ ahora.ToString("yyyyMMdd_HHmmss") + ".xlsx";

                    ExcelComparativaCurva(save.FileName, row.Cells[0].Value.ToString(),
                    Convert.ToDateTime(row.Cells[3].Value),
                    Convert.ToDateTime(row.Cells[4].Value),
                    Convert.ToBoolean(row.Cells[6].Value));
                    exportado_algo = true;
                    
                }
                                       
            }

           

            if (exportado_algo)
            {
                MessageBox.Show("Informe terminado."
                    + System.Environment.NewLine
                    + System.Environment.NewLine
                    + "Exportación Finalizada",
                "Exportación Curvas horarias a Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            }
        }
                
        private void ExcelComparativaCurva(string fichero, string cups20, DateTime fd, DateTime fh, bool devolucion_de_energia)
        {

            int f = 0;
            int c = 0;
            double dif = 0;
            double max = 0;
            double min = 0;
            int numPeriodosHorarios = 0;
            DateTime fecha = new DateTime();
            bool firstOnly = true;
            bool firstOnlyhora2 = true;

            // List<EndesaEntity.medida.CurvaDeCarga> lista_sce = new List<EndesaEntity.medida.CurvaDeCarga>();
            //List<EndesaEntity.medida.CurvaDeCarga> lista_adif = new List<EndesaEntity.medida.CurvaDeCarga>();
            List<EndesaEntity.CurvaCuartoHoraria> lista_adif = new List<EndesaEntity.CurvaCuartoHoraria>();

            List<EndesaEntity.CurvaCuartoHoraria> lista_bi = new List<EndesaEntity.CurvaCuartoHoraria>();
            

            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#E7B6DD");            

            try
            {
                

                FileInfo fileInfo = new FileInfo(fichero);
                if (fileInfo.Exists)
                    fileInfo.Delete();


                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add(cups20);

                var headerCells = workSheet.Cells[3, 1, 3, 7];
                var headerFont = headerCells.Style.Font;
                headerFont.Bold = true;

                headerCells = workSheet.Cells["C1:F1"];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;

                headerCells = workSheet.Cells["C2:I2"];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;

                f = 1;
                c = 1;

                // TextoCentrado(excelPackage, 3, f, c);
                workSheet.Cells[f, 3].Value = "BI"; workSheet.Cells["C1:D1"].Merge = true; c++;
                workSheet.Cells[f, 5].Value = "ADIF"; workSheet.Cells["E1:F1"].Merge = true; c++;
                f++;
                workSheet.Cells[f, 8].Value = "MAX DIF"; c++;
                workSheet.Cells[f, 9].Value = "MIN DIF"; c++;
                f++;
                c = 1;
                workSheet.Cells[f, c].Value = "Fecha"; c++;
                workSheet.Cells[f, c].Value = "Hora"; c++;
                workSheet.Cells[f, c].Value = "Activa"; c++;
                workSheet.Cells[f, c].Value = "Reactiva"; c++;
                workSheet.Cells[f, c].Value = "Activa"; c++;
                workSheet.Cells[f, c].Value = "Reactiva"; c++;
                workSheet.Cells[f, c].Value = "DIF_ACTIV"; c++;
                

                for (int i = 1; i <= c; i++)
                {
                    // xls.PonEstilo(f, i, office.Excel.Estilos.NEGRITA);
                }
                

                EndesaEntity.medida.PuntoSuministro punto_suministro;
                if (pm.dic.TryGetValue(cups20, out punto_suministro))
                {
                    // lista_sce = cc_sce.GetCurva(punto_suministro.cups13, fd, fh);
                                      

                    lista_bi = cc_bi.GetCurvaCups20(punto_suministro.cups20, fd, fh);
                    //lista_adif = medida_adif.GetCurva(punto_suministro.cups20, fd, fh);
                    

                }

                // TOTALES
                //workSheet.Cells[2, 3].Value = cc_sce.TotalActiva(punto_suministro.cups13, fd, fh); workSheet.Cells[2, 3].Style.Numberformat.Format = "#,##0";
                //workSheet.Cells[2, 4].Value = cc_sce.TotalReactiva(punto_suministro.cups13, fd, fh); workSheet.Cells[2, 4].Style.Numberformat.Format = "#,##0";

                workSheet.Cells[2, 3].Value = cc_bi.TotalActivaCups20(punto_suministro.cups20, fd, fh); workSheet.Cells[2, 3].Style.Numberformat.Format = "#,##0";
                workSheet.Cells[2, 4].Value = cc_bi.TotalReactivaCups20(punto_suministro.cups20, fd, fh); workSheet.Cells[2, 4].Style.Numberformat.Format = "#,##0";

                if (devolucion_de_energia)
                    workSheet.Cells[2, 5].Value = medida_adif.Neteo(punto_suministro.cups20, fd, fh); 
                else
                    workSheet.Cells[2, 5].Value = medida_adif.TotalActivaEntrante(punto_suministro.cups20, fd, fh);

                workSheet.Cells[2, 5].Style.Numberformat.Format = "#,##0";

                workSheet.Cells[2, 6].Value = medida_adif.TotalReactivaR4(punto_suministro.cups20, fd, fh); workSheet.Cells[2, 6].Style.Numberformat.Format = "#,##0";


                //for (int i = 0; i < lista_sce.Count(); i++)
                for (int i = 0; i < lista_bi.Count(); i++)
                {
                    //numPeriodosHorarios = NumPeriodosHorarios(lista_sce[i].fecha);
                    numPeriodosHorarios = NumPeriodosHorarios(lista_bi[i].fecha);
                    //fecha = lista_sce[i].fecha;
                    fecha = lista_bi[i].fecha;


                    if (devolucion_de_energia)
                        lista_adif = medida_adif.GetCurvaPO1011_Neteada(punto_suministro.cups20, fecha, fecha);
                    else
                        lista_adif = medida_adif.GetCurvaPO1011(punto_suministro.cups20, fecha, fecha);

                    for (int j = 1; j <= numPeriodosHorarios; j++)
                    {
                        f++;
                        c = 1;
                        //workSheet.Cells[f, c].Value = lista_sce[i].fecha; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                        workSheet.Cells[f, c].Value = lista_bi[i].fecha; 
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;


                        if (numPeriodosHorarios == 23 && fecha.Hour >= 2)
                        {
                            if (firstOnly)
                            {
                                fecha = fecha.AddHours(1);
                                firstOnly = false;
                            }

                            workSheet.Cells[f, c].Value = fecha.ToString("HH:mm"); c++;
                            fecha = fecha.AddHours(1);                            
                            
                            workSheet.Cells[f, c].Value = lista_bi[i].a[j + 1]; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                            workSheet.Cells[f, c].Value = lista_bi[i].r1[j + 1]; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                            if (devolucion_de_energia)
                                workSheet.Cells[f, c].Value = lista_adif[j-1].NETEADA; 
                            else
                                workSheet.Cells[f, c].Value = lista_adif[j - 1].AE;

                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                            c++;

                            workSheet.Cells[f, c].Value = lista_adif[j-1].R4; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                            if (devolucion_de_energia)
                                dif = (lista_bi[i].a[j + 1] - lista_adif[j-1].NETEADA);
                            else
                                dif = (lista_bi[i].a[j + 1] - lista_adif[j - 1].AE);

                        }
                        else if (numPeriodosHorarios == 25)
                        {
                            if(fecha.Hour <= 2)
                            {
                                workSheet.Cells[f, c].Value = fecha.ToString("HH:mm"); c++;
                                fecha = fecha.AddHours(1);

                                workSheet.Cells[f, c].Value = lista_bi[i].a[j]; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                workSheet.Cells[f, c].Value = lista_bi[i].r1[j]; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                                if (devolucion_de_energia)
                                    workSheet.Cells[f, c].Value = lista_adif[j - 1].NETEADA; 
                                else
                                    workSheet.Cells[f, c].Value = lista_adif[j - 1].AE;

                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; 
                                c++;



                                workSheet.Cells[f, c].Value = lista_adif[j - 1].R4; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                                if (devolucion_de_energia)
                                    dif = (lista_bi[i].a[j] - lista_adif[j - 1].NETEADA);
                                else
                                    dif = (lista_bi[i].a[j] - lista_adif[j - 1].AE);
                            }
                            else
                            {
                                workSheet.Cells[f, c].Value = fecha.AddHours(-1).ToString("HH:mm"); c++;
                                fecha = fecha.AddHours(1);

                                workSheet.Cells[f, c].Value = lista_bi[i].a[j]; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                workSheet.Cells[f, c].Value = lista_bi[i].r1[j]; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                                if (devolucion_de_energia)
                                    workSheet.Cells[f, c].Value = lista_adif[j - 1].NETEADA; 
                                else
                                    workSheet.Cells[f, c].Value = lista_adif[j - 1].AE;


                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                workSheet.Cells[f, c].Value = lista_adif[j - 1].R4; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                                if (devolucion_de_energia)
                                    dif = (lista_bi[i].a[j] - lista_adif[j - 1].NETEADA);
                                else
                                    dif = (lista_bi[i].a[j] - lista_adif[j - 1].AE);
                            }
                            
                        }
                        else
                        {
                            workSheet.Cells[f, c].Value = fecha.ToString("HH:mm"); c++;
                            fecha = fecha.AddHours(1);
                            
                            workSheet.Cells[f, c].Value = lista_bi[i].a[j]; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                            workSheet.Cells[f, c].Value = lista_bi[i].r1[j]; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                            if (devolucion_de_energia)
                                workSheet.Cells[f, c].Value = lista_adif[j - 1].NETEADA; 
                            else
                                workSheet.Cells[f, c].Value = lista_adif[j - 1].AE;

                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                            workSheet.Cells[f, c].Value = lista_adif[j - 1].R4; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                            if (devolucion_de_energia)
                                dif = (lista_bi[i].a[j] - lista_adif[j - 1].NETEADA);
                            else
                                dif = (lista_bi[i].a[j] - lista_adif[j - 1].AE);
                        }    

                        
                        if (dif != 0)
                        {
                            if (dif > max)
                                max = dif;
                            if (dif < min)
                                min = dif;

                            for(int w = 1; w < 8; w++)
                            {
                                workSheet.Cells[f, w].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[f, w].Style.Fill.BackgroundColor.SetColor(colFromHex);
                            }
                            
                        }
                        workSheet.Cells[f, c].Value = dif; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    }


                }

                workSheet.Cells[3, 8].Value = max; workSheet.Cells[3, 8].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[3, 9].Value = min; workSheet.Cells[3, 9].Style.Numberformat.Format = "#,##0"; c++;

                var allCells = workSheet.Cells[1, 1, f, c];

                allCells.AutoFitColumns();

                workSheet.View.FreezePanes(4, 1);
                workSheet.Cells["A3:G3"].AutoFilter = true;
                allCells.AutoFitColumns();

                excelPackage.Save();

           

            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,              
            "Error en Exportación Curvas cuarto-horarias a Excel",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
            }
        }

        private void TextoCentrado(ExcelPackage excelPackage, int hoja, int fila, int columna)
        {
            var workSheet = excelPackage.Workbook.Worksheets[hoja];
            workSheet.Cells[fila, columna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        private int NumPeriodosHorarios(DateTime f)
        {
            return (f.Month == 3 ?
                    f.Equals(UltimoDomingoMarzo(f)) ? 23 : 24 :
                    f.Month == 10 ?
                    f.Equals(UltimoDomingoOctubre(f)) ? 25 : 24 : 24);
        }

        private DateTime UltimoDomingoMarzo(DateTime f)
        {
            return new DateTime(f.Year, 3, 31).AddDays(-(new DateTime(f.Year, 3, 31).DayOfWeek - DayOfWeek.Sunday));
        }

        private DateTime UltimoDomingoOctubre(DateTime f)
        {
            return new DateTime(f.Year, 10, 31).AddDays(-(new DateTime(f.Year, 10, 31).DayOfWeek - DayOfWeek.Sunday));
        }

        private void txtFFACTDES_Leave(object sender, EventArgs e)
        {
            this.InitListBox(Convert.ToDateTime(txtFFACTDES.Value), Convert.ToDateTime(txtFFACTHAS.Value));
        }

        private void txtFFACTHAS_Leave(object sender, EventArgs e)
        {
            this.InitListBox(Convert.ToDateTime(txtFFACTDES.Value), Convert.ToDateTime(txtFFACTHAS.Value));
        }

        private void curvaResumenToolStripMenuItem_Click(object sender, EventArgs e)
        {            

            foreach (DataGridViewRow row in dgv.SelectedRows)
            {
                forms.medida.FrmCurvaResumen f = new medida.FrmCurvaResumen(row.Cells[0].Value.ToString(), 
                    Convert.ToDateTime(row.Cells[3].Value), Convert.ToDateTime(row.Cells[4].Value));

                f.Show();

            }

            


        }

        private void FrmAdif_CS_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.Start("Medida", "FrmAdif_CS" ,"N/A");
        }

        private void cargaFicherosADIFToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            int totalArchivos = 0;
            double percent = 0;
            forms.FrmProgressBar pb = new forms.FrmProgressBar();
            int i = 0;
            int ii = 0;
            DateTime begin = new DateTime();
            int total_cups = 0;

            EndesaBusiness.adif.AdifImportar adif_imp = new EndesaBusiness.adif.AdifImportar();

            begin = DateTime.Now;
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "ficheros formato PO.10.11|*.ZIP";
            d.Multiselect = true;
            EndesaBusiness.adif.P01011_Funciones pp = new EndesaBusiness.adif.P01011_Funciones();
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                if (!Directory.Exists(p.GetValue("inbox")))
                    Directory.CreateDirectory(p.GetValue("inbox"));

                BorrarContenidoDirectorio(p.GetValue("inbox", DateTime.Now, DateTime.Now));

                totalArchivos = d.FileNames.Count();

                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = 100;
                pb.Text = "Descomprimiendo ...";

                foreach(string fileName in d.FileNames)
                {
                    ii++;
                    percent = (ii / Convert.ToDouble(d.FileNames.Count())) * 100;
                    pb.progressBar.Value = Convert.ToInt32(percent);
                    pb.progressBar.Value = Convert.ToInt32(percent) - 1;
                    pb.txtDescripcion.Text = "Extrayendo " + ii.ToString("#,##0") + " / " + d.FileNames.Count().ToString("#,##0") + " archivos --> " + fileName;
                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    pb.Refresh();

                    EndesaBusiness.utilidades.ZipUnZip zip = new EndesaBusiness.utilidades.ZipUnZip();
                    zip.Descomprimir(fileName, p.GetValue("inbox"));
                    
                }

                pb.Close();

                DirectoryInfo DirInfo = new DirectoryInfo(p.GetValue("inbox"));
                var files = from f in DirInfo.EnumerateFiles("*CC1_*", SearchOption.AllDirectories) select f;

                pb = new forms.FrmProgressBar();
                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = 100;
                pb.Text = "Descomprimiendo ...";

                ii = 0;

                foreach (var file in files)
                {
                    i++;
                    if (i == 100)
                    {
                        i = 0;
                        pp.Save();
                        pp = null;
                        pp = new EndesaBusiness.adif.P01011_Funciones();
                    }



                    ii++;
                    percent = (ii / Convert.ToDouble(files.Count())) * 100;
                    pb.progressBar.Value = Convert.ToInt32(percent);
                    pb.progressBar.Increment(-1);
                    pb.txtDescripcion.Text = "Importando " + ii.ToString("#,##0") + " / " + files.Count().ToString("#,##0") + " archivos --> " + file.Name;
                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    pb.Refresh();

                    pp.CargaP01011(file.FullName);
                    pp.GuardaP01011(file.FullName);

                }

                pb.Close();
                pp.Save();

                total_cups = adif_imp.TratarCarga();
                pf.SaveProcess("Importar PO1011", "Importacion curvas ADIF", begin, DateTime.Now);

                MessageBox.Show("La importación ha concluido correctamente."
                    + System.Environment.NewLine + System.Environment.NewLine
                    + "Ficheros Seleccionados: " + totalArchivos + System.Environment.NewLine
                    + "Ficheros cargados: " + totalArchivos + System.Environment.NewLine
                    + "CUPS cargados: " + total_cups,
                    "Importación ficheros ficheros formato PO.10.11 ADIF",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }

        }
        private void cargaFicherosADIFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int totalArchivos = 0;
            forms.FrmProgressBar pb = new forms.FrmProgressBar();
            int i = 0;
            DateTime begin = new DateTime();
            int total_cups = 0;

            EndesaBusiness.adif.AdifImportar adif_imp = new EndesaBusiness.adif.AdifImportar();

            begin = DateTime.Now;
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "ficheros formato PO.10.11|*CC1_*";
            d.Multiselect = true;
            EndesaBusiness.adif.P01011_Funciones pp = new EndesaBusiness.adif.P01011_Funciones();
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                totalArchivos = d.FileNames.Count(); 

                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = d.FileNames.Count();
                pb.Text = "Importando... ";

                foreach (string fileName in d.FileNames)
                {
                    i++;
                    if (i == 100)
                    {
                        i = 0;
                        pp.Save();
                        pp = null;
                        pp = new EndesaBusiness.adif.P01011_Funciones();
                    }

                    pb.txtDescripcion.Text = "Importando: " + fileName.ToString();
                    pb.progressBar.Increment(1);                    
                    pb.Refresh();

                    pp.CargaP01011(fileName);
                    pp.GuardaP01011(fileName);

                }

                pb.Close();
                pp.Save();

                total_cups = adif_imp.TratarCarga();
                pf.SaveProcess("Importar PO1011", "Importacion curvas ADIF", begin, DateTime.Now);

                MessageBox.Show("La importación ha concluido correctamente."
                    + System.Environment.NewLine + System.Environment.NewLine
                    + "Ficheros Seleccionados: " + totalArchivos + System.Environment.NewLine
                    + "Ficheros cargados: " + totalArchivos + System.Environment.NewLine
                    + "CUPS cargados: " + total_cups,
                    "Importación ficheros ficheros formato PO.10.11 ADIF",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private void BorrarContenidoDirectorio(string directorio)
        {
            string[] listaArchivos;
            FileInfo file;

            listaArchivos = Directory.GetFiles(directorio);
            for (int i = 0; i < listaArchivos.Count(); i++)
            {
                file = new FileInfo(listaArchivos[i]);
                file.Delete();
            }
        }
    }
}
