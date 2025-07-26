using EndesaEntity.eer;
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
    public partial class FrmFacturadorEER : Form
    {
        EndesaBusiness.eer.Inventario inventario;
        EndesaBusiness.eer.FacturasEER facturas;
        EndesaBusiness.medida.CurvaResumenEER cr;
        EndesaBusiness.utilidades.Param p;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmFacturadorEER()
        {
            usage.Start("Facturación", "FrmFacturadorEER" ,"N/A");
            InitializeComponent();
        }

        private void Facturador_Load(object sender, EventArgs e)
        {
            p = new EndesaBusiness.utilidades.Param("eer_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);

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
            cr = new EndesaBusiness.medida.CurvaResumenEER(fd, fh);
            facturas = new EndesaBusiness.eer.FacturasEER(fd, fh);
            inventario = new EndesaBusiness.eer.Inventario(fd, fh);
            EndesaBusiness.medida.DatosPeajesEER peajesEER = 
                new EndesaBusiness.medida.DatosPeajesEER(fd, fh);

            foreach (EndesaEntity.eer.InventarioPuntosFacturador p in inventario.lista_inventario)
            {
                facturas.GetInvoice(p.cups20, p.fecha_consumo_desde, p.fecha_consumo_hasta);
                if (facturas.existe)
                {
                    p.estado = facturas.estado;
                    p.num_factura = facturas.codigo_factura;
                    p.fecha_factura = facturas.fecha_factura;
                }
                else if (cr.CurvaCompleta(p.cups20, p.fecha_consumo_desde, p.fecha_consumo_hasta) || peajesEER.HayDatos(p.cups20))
                    p.estado = "Generar Prefactura";
                else
                    p.estado = "Medida Incompleta";
                    

            }

            dgv_Inventario.AutoGenerateColumns = false;           


            dgv_Inventario.DataSource = inventario.lista_inventario;
            lbl_registros_i.Text = string.Format("Registros: {0:#,##0}", inventario.lista_inventario.Count());

            foreach (DataGridViewRow row in dgv_Inventario.Rows)
                if ((row.Cells[6].Value.ToString()) == "Medida Incompleta")
                    row.DefaultCellStyle.BackColor = Color.LightPink;
            else if((row.Cells[6].Value.ToString()) == "Facturada")
                    row.DefaultCellStyle.BackColor = Color.LightGreen;
            else if ((row.Cells[6].Value.ToString()) == "Calculada")
                    row.DefaultCellStyle.BackColor = Color.Yellow;

            Cursor.Current = Cursors.Default;
        }

        private void FacturarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<EndesaEntity.punto_suministro.PuntoSuministro> lista_PS =
                new List<EndesaEntity.punto_suministro.PuntoSuministro>();

            DateTime fd = new DateTime();
            DateTime fh = new DateTime();

            fd = Convert.ToDateTime(txt_fecha_consumo_desde.Value);
            fh = Convert.ToDateTime(txt_fecha_consumo_hasta.Value);

            EndesaBusiness.eer.Facturador facturador;


            foreach (DataGridViewRow row in dgv_Inventario.SelectedRows)
            {

                DataGridViewCell cups = (DataGridViewCell)
                 row.Cells[2];

                DataGridViewCell fecha_consumo_desde = (DataGridViewCell)
                    row.Cells[4];

                DataGridViewCell fecha_consumo_hasta = (DataGridViewCell)
                    row.Cells[5];

                foreach (EndesaEntity.eer.InventarioPuntosFacturador p in inventario.lista_inventario)
                {
                    if (p.cups20 == cups.Value.ToString() &&
                        (p.fecha_consumo_desde == Convert.ToDateTime(fecha_consumo_desde.Value.ToString()) && 
                            p.fecha_consumo_hasta == Convert.ToDateTime(fecha_consumo_hasta.Value.ToString())))
                        lista_PS.Add(inventario.GetPS(p.cups20, p.fecha_consumo_desde, p.fecha_consumo_hasta));
                }
            }

            if (lista_PS.Count() > 0)
            {
                DialogResult result3 = MessageBox.Show("¿Desea lanzar la facturación para " + lista_PS.Count() + " puntos",
                "Endesa Energía Renovables",
                            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result3 == DialogResult.Yes)
                {
                    facturador = new EndesaBusiness.eer.Facturador(fd, fh, lista_PS, inventario, chk_exportar_calculos.Checked);
                    MessageBox.Show("Facturación completada.",
                        "Endesa Energía Renovables",
                           MessageBoxButtons.OK, MessageBoxIcon.Information);


                    CargaDGVs(fd, fh);
                }
            }

        }

        private void BtnRefresh_Click(object sender, EventArgs e)
        {

            DateTime fd = new DateTime();
            DateTime fh = new DateTime();

            fd = Convert.ToDateTime(txt_fecha_consumo_desde.Value);
            fh = Convert.ToDateTime(txt_fecha_consumo_hasta.Value);

            CargaDGVs(fd, fh);
        }

        private void ImportarCurvasDeAccessAMySQLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.eer.CurvasAccess ccAccess = new EndesaBusiness.eer.CurvasAccess();
            ccAccess.Proceso();
        }

        private void generarPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void parámetrosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new FrmParameters();
            p.tabla = "eer_param";
            p.esquemaString = "CON";
            p.Show();
        }

        private void dgv_Inventario_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
                                 

            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv_Inventario.CurrentRow.Index;
            int c = 0;

            FrmFacturadorInvoiceNumberEdit f = new FrmFacturadorInvoiceNumberEdit();
            f.Text = "Registrar Factura";

                   

            row = dgv_Inventario.Rows[fila];

            if (row.Cells[c].Value != null)
                f.txt_cnifdnic.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_dapersoc.Text = row.Cells[c].Value.ToString(); c++;
                        

            if (row.Cells[c].Value != null)
                f.txt_cups20.Text = row.Cells[c].Value.ToString(); c++;


            if (row.Cells[4].Value != null)
                f.txt_fecha_consumo_desde.Value = Convert.ToDateTime(row.Cells[4].Value);

            if (row.Cells[5].Value != null)
                f.txt_fecha_consumo_hasta.Value = Convert.ToDateTime(row.Cells[5].Value);


            f.ShowDialog();           

            CargaDGVs(txt_fecha_consumo_desde.Value, txt_fecha_consumo_hasta.Value);


        }

        private void informeDeFacturasAExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {

            forms.facturacion.FrmInformeEERcs f = new FrmInformeEERcs();
            f.ShowDialog();
            
            
        }

        private void cerrarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dgv_Inventario_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void facturasManualesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmEER_FacturasManuales f = new FrmEER_FacturasManuales();
            f.Show();
        }

        private void FrmFacturadorEER_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmFacturadorEER" ,"N/A");
        }
    }
}
