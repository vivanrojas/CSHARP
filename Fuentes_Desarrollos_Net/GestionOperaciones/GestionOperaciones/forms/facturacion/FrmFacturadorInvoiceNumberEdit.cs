// using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
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
    public partial class FrmFacturadorInvoiceNumberEdit : Form
    {
        

        public FrmFacturadorInvoiceNumberEdit()
        {
            InitializeComponent();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (txtFechaFactura.Text == null || txtFechaFactura.Text == "")
                errorProvider.SetError(txtFechaFactura, "El campo debe estar informado.");
            else if (txt_numFactura.Text == null || txt_numFactura.Text == "")
                errorProvider.SetError(txt_numFactura, "El campo debe estar informado.");            
            else
            {

                EndesaBusiness.eer.FacturasEER factura = new EndesaBusiness.eer.FacturasEER();
                factura.RegistraFactura(txt_cups20.Text,
                    txt_fecha_consumo_desde.Value, txt_fecha_consumo_hasta.Value,
                    txt_numFactura.Text, txtFechaFactura.Value);
                
                this.Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
