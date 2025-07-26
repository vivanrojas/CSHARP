using OfficeOpenXml;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmFacturadorPortugal : Form
    {
        EndesaBusiness.facturacion.InventarioFacturacionPortugal inventario;
        EndesaBusiness.facturacion.Clicks clicks;
        EndesaBusiness.facturacion.Spot spot;
        List<EndesaEntity.facturacion.InventarioFacturacion> lista_inventario;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        EndesaBusiness.utilidades.Param p;
        public FrmFacturadorPortugal()
        {
            usage.Start("Facturación", "FrmFacturadorPortugal" ,"N/A");
            InitializeComponent();
        }

        private void FrmFacturadorPortugal_Load(object sender, EventArgs e)
        {

                        

            p = new EndesaBusiness.utilidades.Param("ag_pt_param", EndesaBusiness.servidores.MySQLDB.Esquemas.FAC);

            DateTime mesAnterior = DateTime.Now.AddMonths(-1);
            //DateTime mesAnterior = DateTime.Now;
            int anio = mesAnterior.Year;
            int mes = mesAnterior.Month;
            int dias_del_mes = DateTime.DaysInMonth(anio, mes);

            DateTime fh = new DateTime(anio, mes, dias_del_mes);
            DateTime fd = new DateTime(fh.Year, fh.Month, 1);

            txt_fecha_consumo_desde.Value = fd;
            txt_fecha_consumo_hasta.Value = fh;

            //txtNif.Text = p.GetValue("NIF_DIA", DateTime.Now, DateTime.Now);

            CargaDGVs(fd, fh);
            
        }

        private void CargaDGVs(DateTime fd, DateTime fh)
        {
            
            Cursor.Current = Cursors.WaitCursor;
            inventario = new EndesaBusiness.facturacion.InventarioFacturacionPortugal(txtNif.Text == "" ? null : txtNif.Text, fd, fh);
            clicks = new EndesaBusiness.facturacion.Clicks(fd, fh);
            spot = new EndesaBusiness.facturacion.Spot(fd, fh);

            #region Inventario
            dgv_Inventario.AutoGenerateColumns = false;
            //lista_inventario = inventario.dic_inventario.Select(z => z.Value).Where(z => z.actualizado == false).ToList();
            lista_inventario = inventario.dic_inventario.Select(z => z.Value).ToList();
            dgv_Inventario.DataSource = lista_inventario;
            lbl_registros_i.Text = string.Format("Registros: {0:#,##0}", lista_inventario.Count());

            foreach (DataGridViewRow row in dgv_Inventario.Rows)
            {
                if(row.Cells[4].Value.ToString() != "FACTURADO")
                {
                    if (Convert.ToInt32((row.Cells[4].Value).ToString().Substring(0, 2)) < 8)
                        row.DefaultCellStyle.BackColor = Color.Tomato;                    
                }
                else
                    row.DefaultCellStyle.BackColor = Color.LightSeaGreen;


            }

                
            #endregion


            dgv_Spot.AutoGenerateColumns = false;
            dgv_Spot.DataSource = spot.lista;
            lbl_precio_medio.Text = string.Format("Precio medio: {0:#,##0.##} € MW", spot.precio_medio);
            lbl_registros_s.Text = string.Format("Registros: {0:#,##0}", spot.lista.Count());

            dgv_Clicks.AutoGenerateColumns = false;
            dgv_Clicks.DataSource = clicks.dic.Select(z => z.Value).ToList();
            Cursor.Current = Cursors.Default;

        }
        private void btnFacturar_Click(object sender, EventArgs e)
        {

            DateTime fd = new DateTime();
            DateTime fh = new DateTime();
            int total_puntos_a_facturar = 0;
            List<string> lista_cups13_a_facturar = new List<string>();
            List<string> lista_cups20_a_facturar = new List<string>();
            EndesaBusiness.facturacion.InventarioFacturacionPortugal inventarioDIA;
            EndesaBusiness.facturacion.PagosPortugal pp;

            EndesaBusiness.facturacion.FacturaPortugalExcel excel = new EndesaBusiness.facturacion.FacturaPortugalExcel();

            try
            {
                fd = Convert.ToDateTime(txt_fecha_consumo_desde.Text);
                fh = Convert.ToDateTime(txt_fecha_consumo_hasta.Text);

                inventarioDIA = new EndesaBusiness.facturacion.InventarioFacturacionPortugal(
                    p.GetValue("NIF_DIA", DateTime.Now, DateTime.Now),fd, fh);

                foreach (KeyValuePair<string, EndesaEntity.facturacion.InventarioFacturacion> p in inventario.dic_inventario)
                {
                    if (!p.Value.actualizado)
                    {
                        if (p.Value.ltp != "FACTURADO")
                        {
                            if (Convert.ToInt32(p.Value.ltp.Substring(0, 2)) == 10)
                            {
                                total_puntos_a_facturar++;
                                lista_cups13_a_facturar.Add(p.Value.cups13);
                                lista_cups20_a_facturar.Add(p.Key);
                            }
                        }

                    }

                }

                pp = new EndesaBusiness.facturacion.PagosPortugal(lista_cups20_a_facturar, fd, fh);


                if (!spot.hayDatos)
                {
                    MessageBox.Show("El número de precios Spot cargados son " + String.Format("{0:#,##0}", spot.lista.Count())
                        + " cuando deberían ser " + String.Format("{0:#,##0}", spot.totalPeriodosCuartoHorarios) + "."
                        + System.Environment.NewLine
                        + "No se puede realizar la facturación.",
                     "Facturador Portugal",
                     MessageBoxButtons.OK,
                   MessageBoxIcon.Information);

                }
                else if (total_puntos_a_facturar > 0)
                {
                    DialogResult result3 = MessageBox.Show("¿Desea lanzar la facturación para " + total_puntos_a_facturar + " puntos",
                    "Facturador de Portugal",
                                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    if (result3 == DialogResult.Yes)
                    {
                        //dirPS = new GO.facturacion.portugal.DireccionPS(inventario.dic_inventario.Select(z => z.Key).ToList());

                        // Cargamos medida
                        EndesaBusiness.medida.CCRD medida_cc =
                            new EndesaBusiness.medida.CCRD(lista_cups13_a_facturar, fd, fh);

                        excel.GeneraExcelDIA(inventario, pp, medida_cc, fd, fh);


                        //foreach (KeyValuePair<string, EndesaEntity.facturacion.InventarioFacturacion> p in inventario.dic_inventario)
                        //{

                        //    if (medida_cc.CurvaCompleta(p.Value.cups13))
                        //    {
                                
                        //        inventario.cpe = p.Value.cpe;
                        //        inventario.estado = "facturador generado";
                        //        inventario.actualizado = true;
                        //        inventario.Update();
                        //    }

                        //    else
                        //    {
                        //        inventario.cpe = p.Value.cpe;
                        //        inventario.estado = "CC Incompleta";
                        //        inventario.Update();
                        //    }

                        //}


                    }

                    MessageBox.Show("Proceso finalizado correctamente.",
                      "Facturadores Portugal",
                      MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                }
                else
                {
                    MessageBox.Show("No hay puntos pendientes de facturar.",
                      "Facturadores Portugal",
                      MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message,
                                          "Facturadores Portugal",
                                          MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
            }


        }

        private void importarPreciosSpotToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Ficheros Excel|*.xlsx";
            d.Multiselect = true;
            bool hayError = false;
            EndesaBusiness.facturacion.Spot spot = new EndesaBusiness.facturacion.Spot();

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                
                foreach (string fileName in d.FileNames)
                {                    
                    hayError = spot.ImportarPreciosSpot(fileName);
                    if (hayError)
                        break;
                }
                if (!hayError)
                {
                    
                }
                

            }
        }

       
        

        private void importarClicksToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Ficheros Excel|*.xlsx";
            d.Multiselect = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                foreach (string fileName in d.FileNames)
                {
                    clicks.ImportarClicks(fileName);
                }
                Cursor.Current = Cursors.Default;

                MessageBox.Show("Carga finalizada",
                             "Carga clicks",
                             MessageBoxButtons.OK,
                         MessageBoxIcon.Information);

            }
        }

        

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            CargaDGVs(Convert.ToDateTime(txt_fecha_consumo_desde.Value), Convert.ToDateTime(txt_fecha_consumo_hasta.Value));
        }

        private void generarFacturaATravésDePlantillaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.facturacion.FacturaPortugalExcel excel = new EndesaBusiness.facturacion.FacturaPortugalExcel();
            List<string> lista_cups13_a_facturar = new List<string>();
            DateTime fd = new DateTime();
            DateTime fh = new DateTime();

            fd = Convert.ToDateTime(txt_fecha_consumo_desde.Value);
            fh = Convert.ToDateTime(txt_fecha_consumo_hasta.Value);

            forms.FrmProgressBar pb;
            int progreso_factura = 0;
            double percent = 0;
            FileInfo plantilla;

            foreach (DataGridViewRow row in dgv_Inventario.SelectedRows)
            {

                DataGridViewCell cups = (DataGridViewCell)
                 row.Cells[3];
                foreach (KeyValuePair<string, EndesaEntity.facturacion.InventarioFacturacion> p in inventario.dic_inventario)
                {
                    if (p.Value.cpe == cups.Value.ToString())
                        lista_cups13_a_facturar.Add(p.Value.cups13);
                }
            }

            if (!spot.hayDatos)
            {
                MessageBox.Show("El número de precios Spot cargados son " + String.Format("{0:#,##0}", spot.lista.Count())
                    + " cuando deberían ser " + String.Format("{0:#,##0}", spot.totalPeriodosCuartoHorarios) + "."
                    + System.Environment.NewLine
                    + "No se puede realizar la facturación.",
                 "Facturador Portugal",
                 MessageBoxButtons.OK,
               MessageBoxIcon.Information);

            }
            else if (lista_cups13_a_facturar.Count() > 0)
            {
                DialogResult result3 = MessageBox.Show("¿Desea lanzar la facturación para " + lista_cups13_a_facturar.Count() + " puntos",
                "Facturador de Portugal",
                            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result3 == DialogResult.Yes)
                {
                    pb = new FrmProgressBar();
                    pb.Text = "Generando facturadores";
                    pb.Show();
                    pb.progressBar.Step = 1;
                    pb.progressBar.Maximum = lista_cups13_a_facturar.Count();

                    //dirPS = new GO.facturacion.portugal.DireccionPS(inventario.dic_inventario.Select(z => z.Key).ToList());

                    // Cargamos medida
                    EndesaBusiness.medida.CCRD medida_cc =
                        new EndesaBusiness.medida.CCRD(lista_cups13_a_facturar, fd, fh);

                    for (int y = 0; y < lista_cups13_a_facturar.Count; y++)
                    {
                        foreach (KeyValuePair<string, EndesaEntity.facturacion.InventarioFacturacion> p in inventario.dic_inventario)
                        {
                            if(p.Value.cups13 == lista_cups13_a_facturar[y])
                            {
                                if (medida_cc.CurvaCompleta(p.Value.cups13))
                                {
                                    progreso_factura++;
                                    percent = (progreso_factura / Convert.ToDouble(lista_cups13_a_facturar.Count())) * 100;
                                    pb.progressBar.Increment(1);
                                    pb.txtDescripcion.Text = p.Value.cliente + " " + p.Value.cpe;                                        
                                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                                    pb.Refresh();

                                    // Comprobamos que exista la plantilla

                                    plantilla = new FileInfo(p.Value.ruta_plantilla);
                                    if (plantilla.Exists)
                                    {
                                        excel.GeneraExcel(p.Value.carpeta_cliente, p.Value.ruta_plantilla, fd, fh, spot,
                                        medida_cc.GetCurvaVertical(p.Value.cups13), clicks.GetClicks(p.Value.cpe));
                                        inventario.cpe = p.Value.cpe;
                                        inventario.estado = "facturador generado";
                                        inventario.actualizado = true;
                                        inventario.Update();
                                    }
                                    else
                                    {
                                        inventario.cpe = p.Value.cpe;
                                        inventario.error = "Plantilla no encontrada";
                                        inventario.actualizado = true;
                                        inventario.Update();
                                    }

                                }
                                else
                                {
                                    inventario.cpe = p.Value.cpe;
                                    inventario.estado = "CC Incompleta";
                                    inventario.Update();
                                }
                            }
                            

                        }
                    }

                    if(lista_cups13_a_facturar.Count > 0)
                    {
                        pb.Close();
                        MessageBox.Show("Proceso finalizado correctamente.",
                            "Facturadores Portugal",
                            MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                        CargaDGVs(Convert.ToDateTime(txt_fecha_consumo_desde.Value), Convert.ToDateTime(txt_fecha_consumo_hasta.Value));
                    }
                    

                }

                //        }

                //       

                //    }
                //    else
                //    {
                //        MessageBox.Show("No hay puntos pendientes de facturar.",
                //          "Facturadores Portugal",
                //          MessageBoxButtons.OK,
                //        MessageBoxIcon.Information);
                //    }

                //}
                //catch(Exception ee)
                //{
                //    MessageBox.Show(ee.Message,
                //          "Facturadores Portugal",
                //          MessageBoxButtons.OK,
                //        MessageBoxIcon.Error);
                //}
            }

        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {

        }

        private void FrmFacturadorPortugal_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmFacturadorPortugal" ,"N/A");
        }
    }
}
