using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.contratacion
{
    public partial class FrmEndesaEnergia21_AvisoVentas_Manual : Form
    {
        EndesaBusiness.contratacion.eexxi.Aviso_a_COR aviso_cor;
        public FrmEndesaEnergia21_AvisoVentas_Manual()
        {
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
        

            aviso_cor.GeneraMails_Aviso_Ventas(Convert.ToDateTime(txt_fecha_consumo_desde.Value), 
                Convert.ToDateTime(txt_fecha_consumo_hasta.Value));
        }

        private void FrmEndesaEnergia21_AvisoVentas_Manual_Load(object sender, EventArgs e)
        {
            aviso_cor = new EndesaBusiness.contratacion.eexxi.Aviso_a_COR();
        }
    }
}
