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
    public partial class FrmTunelCuadroMando_Edit : Form
    {

        public EndesaBusiness.facturacion.TunelCuadroMando tunel { get; set; }

        public FrmTunelCuadroMando_Edit()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
                        
            tunel.cliente = txt_cliente.Text;
            tunel.formula_antigua = chk_formula_antigua.Checked;
            tunel.fecha_inicio_tunel = txt_fecha_inicio_tunel.Value;
            tunel.fecha_final_tunel = txt_fecha_fin_tunel.Value;
            tunel.Save();
            this.Close();
        }
    }
}
