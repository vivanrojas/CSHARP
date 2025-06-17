using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Procesos
{
    public partial class FrmTimer : Form
    {

        string cola = "";
        int tiempo;
        public FrmTimer(string _cola)
        {
            cola = _cola;
            InitializeComponent();
            tiempo = 30;
        }

        private void FrmTimer_Load(object sender, EventArgs e)
        {
            this.Text = "Cola de procesos de " + cola;            
            timer1.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            tiempo--;
            txt_tiempo.Text = tiempo.ToString() + " segundos";
            if(tiempo == 0)
            {
                timer1.Enabled = false;
                using (Main m = new Main(cola, true))
                {
                    m.Show();
                }
                    
                this.Close();
            }
        }

        private void txt_tiempo_TextChanged(object sender, EventArgs e)
        {

        }

        private void btn_cancelar_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            Main f = new Main(cola, false);
            f.Show(); 
            
        }
    }
}
