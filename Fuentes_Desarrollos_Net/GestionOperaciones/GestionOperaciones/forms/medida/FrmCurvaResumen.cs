using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.medida
{
    public partial class FrmCurvaResumen : Form
    {

        EndesaBusiness.medida.CurvaResumenSCE cr;

        private string ccounips { get; set; }
        private DateTime fd { get; set; }
        private DateTime fh { get; set; }

        public FrmCurvaResumen(string vccounips, DateTime vfd, DateTime vfh)
        {
            InitializeComponent();
            ccounips = vccounips;
            fd = vfd;
            fh = vfh;

            
        }

        private void FrmCurvaResumen_Load(object sender, EventArgs e)
        {
            Cargadgv();
        }

        private void Cargadgv()
        {
            
            try
            {

                dgv.DataSource = cr.GetCurvaDGV(ccounips, fd, fh);
                
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
               "Error a la hora de buscar la curva resumen",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
