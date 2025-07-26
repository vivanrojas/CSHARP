using EndesaBusiness.servidores;

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
    public partial class FrmEER_FacturasManuales : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        EndesaBusiness.eer.Inventario inventario;
        EndesaBusiness.eer.Factura factura;
        public FrmEER_FacturasManuales()
        {
            usage.Start("Facturación", "FrmEER_FacturasManuales" ,"N/A");
            InitializeComponent();
        }

        private void btn_guardar_Click(object sender, EventArgs e)
        {
            EndesaBusiness.eer.FacturasEER facturaEER = new EndesaBusiness.eer.FacturasEER();
            EndesaEntity.eer.Factura factura = new EndesaEntity.eer.Factura();



            if (txt_cnifdnic.Text == null || txt_cnifdnic.Text == "")
                errorProvider.SetError(txt_cnifdnic, "El campo debe estar informado.");
            else if (txt_dapersoc.Text == null || txt_dapersoc.Text == "")
                errorProvider.SetError(txt_dapersoc, "El campo debe estar informado.");
            else if (txt_cups20.Text == null || txt_cups20.Text == "")
                errorProvider.SetError(txt_cups20, "El campo debe estar informado.");
            else if (txt_cups20.Text.Length != 20)
                errorProvider.SetError(txt_cups20, "El CUPS  no tiene 20 caracteres.");
            else if (txt_fecha_consumo_desde.Text == "")
                errorProvider.SetError(txt_fecha_consumo_desde, "El CUPS  no tiene 20 caracteres.");
            else if (txt_fecha_consumo_hasta.Text == "")
                errorProvider.SetError(txt_fecha_consumo_hasta, "El CUPS  no tiene 20 caracteres.");
            else if (txt_fecha_factura.Text == "")
                errorProvider.SetError(txt_fecha_factura, "El CUPS  no tiene 20 caracteres.");
            else
            {
                factura.id_factura = facturaEER.GetLastID();
                factura.nif = txt_cnifdnic.Text;
                factura.nombre_cliente = txt_dapersoc.Text;
                factura.cupsree = txt_cups20.Text;
                factura.fecha_consumo_desde = Convert.ToDateTime(txt_fecha_consumo_desde.Text);
                factura.fecha_consumo_hasta = Convert.ToDateTime(txt_fecha_consumo_hasta.Text);
                factura.codigo_factura = txt_codigo_factura.Text;
                factura.fecha_factura = Convert.ToDateTime(txt_fecha_factura.Text);
                factura.consumo_activa = Convert.ToInt32(txt_activa.Text.ToString().Replace(".",""));
                factura.consumo_reactiva = Convert.ToInt32(txt_reactiva.Text);
                factura.tarifa = txt_tarifa.Text;
                factura.direccion_facturacion = txt_direccion_fiscal.Text;
                factura.direccion_suministro = txt_direccion_suministro.Text;

                factura.base_ise = Convert.ToDouble(CN(txt_base_ise.Text));
                factura.impuesto_electricidad = Convert.ToDouble(CN(txt_impuesto_electricidad.Text));
                factura.base_ise_reducido = Convert.ToDouble(CN(txt_base_ise_reducido.Text));
                factura.impuesto_electricidad_reducido = Convert.ToDouble(CN(txt_impuesto_electricidad_reducido.Text));
                factura.base_imponible = Convert.ToDouble(CN(txt_base_imponible.Text));
                factura.iva = Convert.ToDouble(CN(txt_iva.Text));
                factura.total_factura = Convert.ToDouble(CN(txt_total_factura.Text));
                factura.termino_energia = Convert.ToDouble(CN(txt_termino_energia.Text));
                factura.facturacion_potencia = Convert.ToDouble(CN(txt_facturacion_potencia.Text));
                factura.recargo_excesos = Convert.ToDouble(CN(txt_recargo_excesos.Text));
                factura.recargo_excesos_reactiva = Convert.ToDouble(CN(txt_recargo_reactiva.Text));

                factura.alquiler = Convert.ToDouble(CN(txt_alquiler.Text));


                facturaEER.GuardaFacturaManual(factura);
            }


           
        }

        private void FrmEER_FacturasManuales_Load(object sender, EventArgs e)
        {
            inventario = new EndesaBusiness.eer.Inventario(DateTime.Now.AddYears(-1), DateTime.Now);
            factura = new EndesaBusiness.eer.Factura();
        }

        private void txt_cups20_TextChanged(object sender, EventArgs e)
        {
            //EndesaEntity.Inventario cliente = inventario.GetCliente(txt_cups20.Text, new DateTime(2022,01,01), new DateTime(2022, 12, 31));
            //EndesaEntity.punto_suministro.PuntoSuministro ps = inventario.GetPS(txt_cups20.Text, txt_cups20.Text, new DateTime(2022, 01, 01), new DateTime(2022, 12, 31));

           


            EndesaEntity.eer.Factura factura_cliente = factura.GetDatosCliente(txt_cups20.Text);

            if (factura_cliente != null)
            {
                txt_cnifdnic.Text = factura_cliente.nif;
                txt_dapersoc.Text = factura_cliente.nombre_cliente;
                txt_direccion_suministro.Text = factura_cliente.direccion_suministro;
                txt_tarifa.Text = factura_cliente.tarifa;
            }

            



        }

        private void ResetCampos()
        {
            txt_cups20.Text = string.Empty;
            txt_cups20.Text = string.Empty;
            txt_dapersoc.Text= string.Empty;
            txt_tarifa.Text= string.Empty;
            txt_direccion_fiscal.Text = string.Empty;
            txt_direccion_suministro.Text = string.Empty;    

            txt_codigo_factura.Text = string.Empty;
            

        }


        private string CN(string numero)
        {
            string n = "";
            n = numero.Replace(".", "");            
            return n;
        }

        private void FrmEER_FacturasManuales_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmEER_FacturasManuales" ,"N/A");
        }
    }
}
